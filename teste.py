def calculate_triangle_area(base, height):
    """Calculate the area of a triangle given base and height."""
    area = (base * height) / 2
    return area

# Get user input
base = float(input("Enter the base of the triangle: "))
height = float(input("Enter the height of the triangle: "))

# Calculate and display the area
area = calculate_triangle_area(base, height)
print(f"The area of the triangle is: {area}")