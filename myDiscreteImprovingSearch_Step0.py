VERYBIGNUMBER = 424242424242
NUMLOCATIONS = 5

import matplotlib.pyplot as plt     
import logging                      
import time                         

logger = logging.getLogger(name='tsp-logger')
logging.basicConfig(level=logging.INFO,
                    format='[%(asctime)s] %(message)s',
                    handlers=[logging.FileHandler("tsp.log")])

def euclidean_distance(point1, point2):
    """
        Euclidean distance between two points
        
        Arguments:
            point1, point2: (x,y)-coordinates of two points
            
        Returns:
            shortest distance from point1 to point2 'as the crow flies' 
    """
    x1, y1 = point1
    x2, y2 = point2
    return ((x1 - x2) ** 2 + (y1 - y2) ** 2) ** 0.5

def total_distance(locations, tour):
    """
        Calculates the total distance of a tour along all locations

        Arguments:
            set of locations ({0,1,...,num_locations-1})
            tour (list): a permutation of the set of locations

        Returns:
            sum of all distances between two consecutive locations in tour 
    """
    totalDistance = 0
    for i in range(len(tour)):
        fromLocation = locations[tour[i]]
        toLocation = locations[tour[(i + 1) % len(tour)]]
        totalDistance += euclidean_distance(fromLocation, toLocation)
    
    return totalDistance

def main():
    
    # Define an instance
    # ----------
    num_locations = 5
    coordinates = [(2.5,5),(0,3),(5,3),(1,0),(4,0)]
    locations = coordinates
    distance = {}
    for i in range(num_locations):
        for j in range(num_locations):
             point1 = coordinates[i]
             point2 = coordinates[j]
             key = (point1,point2)
             distance[key] = euclidean_distance(point1,point2)

    logger.info(msg=f"Specific instance with {num_locations} locations")    

    # Generate a poor tour     
    # ----------
    rnd_tour = [0,1,2,3,4]
    total_distance_rnd_tour = total_distance(coordinates,rnd_tour)
    logger.info(msg=f"Random tour has distance {total_distance_rnd_tour:.2f}") 

    # Plot the random solution and the discrete improving search solution
    # ---

    # Separate the x and y values into separate lists
    x = [coord[0] for coord in coordinates]
    y = [coord[1] for coord in coordinates]

    # In Figure 1, we show both the random solution and the NN solution
    plt.figure(1)
    plt.scatter(x, y)
    plt.xlabel('X-axis Label')
    plt.ylabel('Y-axis Label')
 
    for i in range(1, len(coordinates)):
        x1, y1 = coordinates[rnd_tour[i - 1]]
        x2, y2 = coordinates[rnd_tour[i]]
        plt.plot([x1, x2], [y1, y2], '--', color='grey', linewidth=0.8)  
    x1, y1 = coordinates[rnd_tour[len(coordinates)-1]]
    x2, y2 = coordinates[rnd_tour[0]]
    plt.plot([x1, x2], [y1, y2], '--', color='grey', linewidth=0.8)  
 
    plt.title(f'Solutions on {len(coordinates)} nodes; distance: {total_distance_rnd_tour:.2f}')
    plt.show()

if __name__ == "__main__":
    main()
