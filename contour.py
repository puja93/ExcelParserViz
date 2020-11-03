from skimage import measure
import matplotlib.pyplot as plt


def calculate(segmented_array,visualize=False):
    # Find contours at a constant value of 0.8
    contours = measure.find_contours(segmented_array, 1)

    if(visualize):
        print("Total shapes are",len(contours))

        # Display the image and plot all contours found
        fig, ax = plt.subplots(figsize=(9, 9))
        ax.imshow(segmented_array, cmap=plt.cm.gray)

        for n, contour in enumerate(contours):
            ax.plot(contour[:, 1], contour[:, 0], linewidth=1)

        ax.axis('image')
        ax.set_xticks([])
        ax.set_yticks([])
        plt.show()

    return contours

def take_second(elem):
    return elem[1]

def take_first(elem):
    return elem[0]


def get_splitting_points(contour_array):
    result = []
    for cntr in contour_array:
        #dummy = contours[1]
        contour_row_sorted = sorted(cntr,key=take_first,reverse=False)
        contour_col_sorted = sorted(cntr,key=take_second,reverse=False)

        #Somehow x2 always more about 1 or 2 rows
        x1 = contour_row_sorted[0][0]
        x2 = contour_row_sorted[-1][0] - 2
        y1 = contour_col_sorted[0][1]
        y2 = contour_row_sorted[-1][1]

        result.append([x1,x2,y1,y2])
    return result
    #print(x1,x2," ",y1,y2)


    #df_part = df.iloc[int(x1):int(x2),int(y1):int(y2)]
