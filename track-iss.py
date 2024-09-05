import win32com.client
from skyfield.api import Topos, load
from skyfield.sgp4lib import EarthSatellite
import time

# Connect to the ASCOM telescope
telescope = win32com.client.Dispatch("ASCOM.DeviceHub.Telescope")
telescope.Connected = True

# ISS TLE data
line1 = "1 25544U 98067A   20238.48945602  .00001378  00000-0  33204-4 0  9993"
line2 = "2 25544  51.6453  64.9860 0002311 269.5528 234.6341 15.49104470242353"

# Load timescale and satellite data
ts = load.timescale()
satellite = EarthSatellite(line1, line2, "ISS (ZARYA)", ts)

# Define your observing location (latitude and longitude)
observer_location = Topos('47.6062 N', '122.3321 W')  # Example for Seattle

# Define the hour angle limit to avoid meridian collision (e.g., 5 hours past meridian)
HOUR_ANGLE_LIMIT = 5  # In hours (telescope will stop/flip if exceeded)

def calculate_movement_rates(current_alt, current_az, target_alt, target_az):
    """Calculate the movement rates based on the differences in alt/az."""
    # Movement rate calculations based on the difference in position
    alt_diff = target_alt - current_alt
    az_diff = target_az - current_az

    # Scale factors for converting position difference into speed
    ra_speed = az_diff * 0.1  # Adjust this scale factor based on your mount's sensitivity
    dec_speed = alt_diff * 0.1  # Adjust this scale factor based on your mount's sensitivity

    return ra_speed, dec_speed

def get_hour_angle():
    """Calculate the hour angle of the ISS relative to the observer's meridian."""
    # Calculate the current position of the ISS
    t = ts.now()
    astrometric = satellite.at(t)
    ra, dec, distance = astrometric.radec()
    
    # Get local sidereal time at the observer's location
    observer_sidereal_time = ts.gmst + observer_location.longitude.hours
    
    # Calculate the hour angle (HA = LST - RA)
    hour_angle = observer_sidereal_time - ra.hours
    if hour_angle < 0:
        hour_angle += 24  # Wrap around to 0-24 hours

    return hour_angle

# Continuous tracking loop
while True:
    # Get the current time
    t = ts.now()

    # Calculate the current position of the ISS
    difference = satellite - observer_location
    topocentric = difference.at(t)
    alt, az, distance = topocentric.altaz()

    # Check the hour angle to ensure the telescope does not pass the meridian limit
    hour_angle = get_hour_angle()
    if abs(hour_angle) > HOUR_ANGLE_LIMIT:
        print(f"Hour angle {hour_angle:.2f} exceeds the limit. Stopping tracking to avoid collision.")
        telescope.Park()  # Park the telescope or SlewToHome to avoid hitting the meridian
        break  # Exit the loop and stop tracking

    # Check if the ISS is above the horizon (alt > 0 degrees)
    if alt.degrees > 0:
        # Calculate movement rates for RA and Dec axes based on current and target positions
        ra_speed, dec_speed = calculate_movement_rates(0, 0, alt.degrees, az.degrees)  # Replace 0, 0 with mount's current alt/az

        # Move the mount according to the calculated RA and Dec speeds
        telescope.MoveAxis(0, ra_speed)  # Move the RA axis
        telescope.MoveAxis(1, dec_speed)  # Move the Dec axis

        print(f"Tracking ISS at Alt: {alt.degrees:.2f}, Az: {az.degrees:.2f}")
        print(f"RA Speed: {ra_speed:.4f}, Dec Speed: {dec_speed:.4f}")
    else:
        # ISS is below the horizon, go to home position
        print("ISS is below the horizon. Slewing telescope to home position.")
        telescope.Park()  # Park the telescope, sending it to home position
        break  # Exit the loop as ISS is no longer trackable

    # Sleep for a short interval before updating positions (e.g., 1 second)
    time.sleep(1)

    # Stop movement for each axis
    telescope.MoveAxis(0, 0)  # Stop RA movement
    telescope.MoveAxis(1, 0)  # Stop Dec movement
