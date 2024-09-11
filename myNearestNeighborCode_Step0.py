def euclidean_distance(point1,point2):
    x1,y1 = point1
    x2,y2 = point2
    return ((x1-x2) ** 2 + (y1-y2) **2) ** 0.5

def main():

    """Define the traveling salesman instance
        - n locations
        - d_(i,j) distance matrix
        - x_i,y_i x,y-coordinates of locations
    """
    num_locations = 4
    coordinates = [(0,0),(1,0),(0,1),(1,1)]
    distance = {}
    for loc1 in range(num_locations):
        for loc2 in range(num_locations):
            pnt1 = coordinates[loc1]
            pnt2 = coordinates[loc2]
            dist = euclidean_distance(pnt1,pnt2)
            key = (pnt1,pnt2)
            distance[key] = dist


    """Make the feasible tour,
       Calculate the tour length
    """
    tour = [0,1,2,3]
    tour_distance = 0
    for loc in range(len(tour)):
        from_loc = coordinates[tour[loc]]
        if (loc<len(tour)-1):
            to_loc = coordinates[tour[loc+1]]
        else:
            to_loc= coordinates[tour[0]]
        tour_distance += euclidean_distance(from_loc,to_loc)
    print(f"Tour has length {tour_distance:.2f}")     

if __name__ == "__main__":
    main()