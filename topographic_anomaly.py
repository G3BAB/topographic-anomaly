
# Usage:
# Measurement data should be in Excel workbook,
# By default the file should be named "data.xlsx". Work sheet should be named "data"
#
# DEM data table should be in Excel workbook,
# By default file should be named "dem.xlsx", work sheet named "dem".
#
# By default both .xlsx files should be in execution directory
# Resulting output will be saved in a new column in data.xlsx.

import pandas as pd
from math import sqrt, atan2, pi

gravity_constant = 6.67430e-11  # m^3 kg^-1 s^-2
rho = 2.670  # g/m^3

class MeasurementPoint:
    def __init__(self, id, north, east, height):
        self.id = id
        self.north = float(north.replace(',', '.')) if isinstance(north, str) else float(north)
        self.east = float(east.replace(',', '.')) if isinstance(east, str) else float(east)
        self.height = float(height.replace(',', '.')) if isinstance(height, str) else float(height)
        self.connected_refs = []
        self.slice_distances = {i: {'closest': float('inf'), 'farthest': 0, 'heights': []} for i in range(1, 9)}

    def connect_reference_point(self, ref_point, distance, angle):
        slice = int((angle + pi) / (pi / 4)) % 8 + 1

        if distance <= 1200:
            self.connected_refs.append({'ref_point': ref_point, 'distance': distance, 'slice': slice})

            slice_data = self.slice_distances[slice]
            slice_data['closest'] = min(slice_data['closest'], distance) if slice_data['closest'] <= 1200 else distance
            slice_data['farthest'] = max(slice_data['farthest'], distance)


    def distance_to(self, other_point):
        return sqrt((self.north - other_point.north) ** 2 + (self.east - other_point.east) ** 2)
    
    def angle_to(self, other_point):
        return atan2(other_point.east - self.east, other_point.north - self.north)
                
class ReferencePoint:
    def __init__(self, id, north, east, height):
        self.id = id
        self.north = float(north.replace(',', '.')) if isinstance(north, str) else float(north)
        self.east = float(east.replace(',', '.')) if isinstance(east, str) else float(east)
        self.height = float(height.replace(',', '.')) if isinstance(height, str) else float(height)


    def distance_to(self, other_point):
        return sqrt((self.north - other_point.north) ** 2 + (self.east - other_point.east) ** 2)
    
    def angle_to(self, other_point):
        return atan2(other_point.east - self.east, other_point.north - self.north)

class ReferencePoint:
    def __init__(self, id, north, east, height):
        self.id = id
        self.north = float(north.replace(',', '.')) if isinstance(north, str) else float(north)
        self.east = float(east.replace(',', '.')) if isinstance(east, str) else float(east)
        self.height = float(height.replace(',', '.')) if isinstance(height, str) else float(height)

def calculate_correction(measurement_point):
    correction = 0
    for slice_data in measurement_point.slice_distances.values():
        if slice_data['farthest'] > 0:
            R = slice_data['farthest']
            r = slice_data['closest']
            ref_heights = [ref_point['ref_point'].height for ref_point in measurement_point.connected_refs if ref_point['slice'] == slice_data]

            
            if ref_heights:
                avg_ref_height = sum(ref_heights) / len(ref_heights)
                avg_height_difference = measurement_point.height - avg_ref_height
            else:
                avg_height_difference = measurement_point.height

            sqrt_part1 = sqrt(avg_height_difference**2 + 0**2)
            sqrt_part2 = sqrt(avg_height_difference**2 + R**2)

            summand = (R - 0 + sqrt_part1 - sqrt_part2)
            correction += summand
    
    correction *= (2/8) * pi * gravity_constant * rho * 10**5
    return correction

# Function to calculate the correction for a single measurement point with detailed output
def calculate_correction_verbose(measurement_point):
    print(f"Calculating correction for Measurement Point ID: {measurement_point.id}")
    print(f"Measurement Point Height: {measurement_point.height}")
    correction = 0
    for i, slice_data in measurement_point.slice_distances.items():
        print(f"Slice {i}:")
        avg_ref_height = 0
        if slice_data['farthest'] > 0:  # Check if there are reference points in the slice
            R = slice_data['farthest']
            r = slice_data['closest']
            
            ref_heights = [ref_point['ref_point'].height for ref_point in measurement_point.connected_refs if ref_point['slice'] == i]
            print(f"  Reference Point Heights in the Slice: {ref_heights}")
            
            # Calculate average reference point height for the slice
            if ref_heights:
                avg_ref_height = sum(ref_heights) / len(ref_heights)
                avg_height_difference = measurement_point.height - avg_ref_height
            else:
                avg_height_difference = measurement_point.height  # Use measurement point height if no reference points

            print(f"  Closest point distance within 1200m (r): {r}")
            print(f"  Farthest point distance within 1200m (R): {R}")
            print(f"  Average reference point height (avg_ref_height): {avg_ref_height}")
            print(f"  Average height difference (h_avg): {avg_height_difference}")

            sqrt_part1 = sqrt(avg_height_difference**2 + 0**2)
            sqrt_part2 = sqrt(avg_height_difference**2 + R**2)
            print(f"  sqrt_part1: {sqrt_part1}")
            print(f"  sqrt_part2: {sqrt_part2}")

            # Apply the corrected sum formula
            summand = (R - 0 + sqrt_part1 - sqrt_part2)
            print(f"  Summand for slice: {summand}")
            correction += summand
        else:
            print("  No reference points in this slice within 1200m.")
    
    correction *= (2/8) * pi * gravity_constant * rho * 1000
    print(f"Total correction for Measurement Point ID {measurement_point.id}: {correction}\n")
    return correction


# Load data from Excel
file_path_meas = 'dane.xlsx'
file_path_ref = 'nmt.xlsx'
measurement_df = pd.read_excel(file_path_meas, sheet_name='dane')
reference_df = pd.read_excel(file_path_ref, sheet_name='nmt')

# Create instances
measurement_points = [MeasurementPoint(row['OBJECTID_1'], row['NG'], row['EG'], row['H']) for _, row in measurement_df.iterrows()]
reference_points = [ReferencePoint(row['OBJECTID'], row['NCN'], row['NCE'], row['Hnorm']) for _, row in reference_df.iterrows()]

# Find and connect reference points, and update slice distances
for m_point in measurement_points:
    for r_point in reference_points:
        distance = m_point.distance_to(r_point)
        angle = m_point.angle_to(r_point)
        m_point.connect_reference_point(r_point, distance, angle)

# for m_point in measurement_points:
#     correction = calculate_correction(m_point)
#     print(f"Measurement Point ID: {m_point.id}, Gravitational Correction: {correction}")

grav_corrections = [calculate_correction(m_point) for m_point in measurement_points]
measurement_df['grav_cor'] = grav_corrections
measurement_df.to_excel('dane.xlsx', sheet_name='dane', index=False)
reference_df.to_excel('nmt.xlsx', sheet_name='nmt', index=False)
