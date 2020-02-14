from bokeh.plotting import figure, output_file, show
import math
from circle import Circle  # circle class, used for making a list of circle objects
import csv
import random

RADIUS = 100

def create_circle_linestring(sub_coords):
    '''
    Creates a list of linestrings that trace a circle around a centre point.
    'coordinates' object is written in kml format, coordinates are returned as a list of tuples
    in the format [(x1,y1),(x2,y2),(x3,y3),...cont]
    '''
    kml_coordinates = ''  # can be returned and written to the kml file eventually
    linestring_list = []

    for theta in range(629):  # iterate angle from 0 to 2*pi
        x = float(sub_coords[0]) + RADIUS*math.cos(theta/100)
        y = float(sub_coords[1]) + RADIUS*math.sin(theta/100)
        linestring_list.append((x,y))
        kml_coordinates = kml_coordinates + str(x) + ',' + str(y) + ',' + '0 ' # coordinates in kml format
    return linestring_list


def circles_overlapping(circle, other_circle):
    '''
    Checks to see if the two circle objects passed through are within 200km of each other, i.e. overlapping circles.
    Returns True is the circles are overlapping
    '''
    deltaX = abs(circle.centre_coords[0] - other_circle.centre_coords[0])  # abs(x2-x1)
    deltaY = abs(circle.centre_coords[1] - other_circle.centre_coords[1])

    distance_between_circles = math.sqrt(deltaX**2 + deltaY**2)  # pythag 

    return distance_between_circles < (2 * RADIUS)  # 200km *  1/111 to approximately convert km to degrees



def get_theoretical_intercepts(circle, other_circle):
    '''
    Calculation for intercept points
    2nd answer at https://math.stackexchange.com/questions/256100/how-can-i-find-the-points-at-which-two-circles-intersect
    '''
    
    R = math.sqrt((other_circle.centre_coords[0]- circle.centre_coords[0])**2 + (other_circle.centre_coords[1]-circle.centre_coords[1])**2)
    a = 1/2
    b = math.sqrt((2*((RADIUS**2 + RADIUS**2)/(R**2)))-1)

    intercept1 = [a * (circle.centre_coords[0] + other_circle.centre_coords[0]) + (a * (b * (other_circle.centre_coords[1] - circle.centre_coords[1]))), 
                  a * (other_circle.centre_coords[1] + circle.centre_coords[1]) + (a * (b * (circle.centre_coords[0] - other_circle.centre_coords[0])))] # top intercept [x,y]

    intercept2 = [a * (circle.centre_coords[0] + other_circle.centre_coords[0]) - (a * (b * (other_circle.centre_coords[1] - circle.centre_coords[1]))),
                  a * (other_circle.centre_coords[1] + circle.centre_coords[1]) - (a * (b * (circle.centre_coords[0] - other_circle.centre_coords[0])))] # bottom intercept [x,y]

    return intercept1, intercept2  # ([x1,y1], [x2,y2])


def get_closest_coord(coord, collection):
    '''
    Returns the closest value in a list (collection) to the given value (coords)
    '''
    dist = lambda s,d: (s[0]-d[0])**2+(s[1]-d[1])**2
    match = min(collection, key=lambda p: dist(p, coord)) 

    return match


def write_csv(circle):
    '''
    Writes the linestring of the argument circle to its specified csv file
    Used primarily for testing
    '''
    with open(circle.filename,'w') as out:
        csv_out=csv.writer(out)
        csv_out.writerow(['x','y'])
        for row in circle.linestring:
            csv_out.writerow(row)


def circles_main():
    '''
    red, blue, yellow and green circles are used for testing and plotting with boken
    Convention for all lists of coordinates is x = list[0] and y = list[1]
    substations is in the form [([x1,y1], name), ([x2,y2], name), ...]
    '''
    '''
    circles = []
    circle_file = open('substations.csv', 'r')
    circle_file.readline()
    for line in circle_file:
        circle_line = line.split(',')
        file_name = circle_line[0] + '.csv'
        circles.append(Circle(circle_line[0], [float(circle_line[1]), float(circle_line[2])], create_circle_linestring([float(circle_line[1]), float(circle_line[2])]), file_name, random.choice(['red', 'blue', 'yellow', 'green', 'purple', 'orange', 'pink'])))
    '''
    red_circle_centre = [50, -10]
    blue_circle_centre = [-20,29]
    yellow_circle_centre = [300, 350]
    green_circle_centre = [0, 0]

    # test cases for kml circles
    red_circle = Circle('red', red_circle_centre, create_circle_linestring(red_circle_centre), 'red_circle.csv', 'red')
    blue_circle = Circle('blue', blue_circle_centre, create_circle_linestring(blue_circle_centre), 'blue_cirlce.csv', 'blue')
    green_circle = Circle('green', green_circle_centre, create_circle_linestring(green_circle_centre), 'green_circle.csv', 'green')
    yellow_circle = Circle('yellow', yellow_circle_centre, create_circle_linestring(yellow_circle_centre), 'yellow_circle.csv', 'yellow')
    

    circles = [red_circle, blue_circle, green_circle, yellow_circle]  # list of all the circle objects
    
    for circle in circles:  # cycle through the list of circles and check which ones are touching, without checking the same pair twice
        for other_circle in circles:
            if not(other_circle == circle):
                if circles_overlapping(circle, other_circle): # if circles are touching (within 200km of each other)
                    print('{} and {} are overlapping'.format(circle.name, other_circle.name))
                    print(get_theoretical_intercepts(circle, other_circle))

                    # get the index of the closest coordinate in the linestring to the theoretical intercept value
                    left_intercept = circle.linestring.index(get_closest_coord(get_theoretical_intercepts(circle, other_circle)[0], circle.linestring)) 
                    right_intercept = circle.linestring.index(get_closest_coord(get_theoretical_intercepts(circle, other_circle)[1], circle.linestring))
                    
                    # determine which occurs first, leftmost intercept or rightmost intercept
                    if right_intercept > left_intercept:
                        top_iteration = right_intercept
                        lower_iteration = left_intercept
                    else:
                        top_iteration = left_intercept
                        lower_iteration = right_intercept
                    
                    # in the case of two overlapping circles, no more than half of the perimeter should be removed. 
                    if (top_iteration - lower_iteration) > 314:  # check is the amount of circle to the removed is greater than half of the linestring
                        to_keep = []  # keep the greater segment of the linestring in the circle
                        for i in range(lower_iteration, top_iteration):
                            to_keep.append(circle.linestring[i])
                        circle.linestring = []
                        for element in to_keep:
                            circle.linestring.append(element)
                        write_csv(circle)
                        print(lower_iteration, top_iteration) 
                    else:
                        to_remove = []  # remove the smaller segment of the linestring in the circle
                        for i in range(lower_iteration, top_iteration):
                            to_remove.append(circle.linestring[i])
                        for element in to_remove:
                            circle.linestring.remove(element)
                        write_csv(circle)
                        print(lower_iteration, top_iteration)                  
                else:
                    print('{} and {} are NOT overlapping\n'.format(circle.name, other_circle.name))
                
    output_file('circle.html')

    # Add plot
    p = figure(  # plotting figure used for testing
        title = 'Circles Testing',
        x_axis_label = 'X axis (km)', 
        y_axis_label = 'Y Axis (km)'
    )
    '''
    for circle in circles:
        for coord in circle.linestring:
            circle.x.append(coord[0])
            circle.y.append(coord[1])
        p.line(circle.x, circle.y, legend=circle.name, line_width=2, color=circle.color)
    # LineString circles
    '''
    
    #Render glyph
    # had to manually format the linestrings to plot them with bokeh
    # an iterative method would be better for dealing with large amounts of circles
    redX = []
    redY = []
    for coord in red_circle.linestring:
        redX.append(coord[0])
        redY.append(coord[1])

    
    blueX = []
    blueY = []
    for coord in blue_circle.linestring:
        blueX.append(coord[0])
        blueY.append(coord[1])
    
    greenX = []
    greenY = []
    for coord in green_circle.linestring:
        greenX.append(coord[0])
        greenY.append(coord[1])

    yellowX = []
    yellowY = []
    for coord in yellow_circle.linestring:
        yellowX.append(coord[0])
        yellowY.append(coord[1])
    
    p.line(redX, redY,legend='Red Circle', line_width=2, color='red')
    p.line(blueX,blueY,legend='Blue Circle', line_width=2, color='blue')
    p.line(greenX,greenY,legend='Green Circle', line_width=2, color='green')
    p.line(yellowX,yellowY,legend='Yellow Circle', line_width=2, color='yellow')
    
    show(p)  # output results in circles.html (opens automatically)
    

if __name__ == '__main__':
    circles_main()