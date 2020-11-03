#!/usr/bin/env python
# coding: utf-8

import numpy as np
import matplotlib.pyplot as plt
from skimage.segmentation import random_walker
from skimage.data import binary_blobs
from skimage.exposure import rescale_intensity
import pandas as pd
import skimage



def show_result(data, markers, labels):
    # Plot results
    fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(9, 6), sharex=True, sharey=True)
    ax1.imshow(data, cmap='gray')
    ax1.axis('off')
    ax1.set_title('Noisy data')
    ax2.imshow(markers, cmap='magma')
    ax2.axis('off')
    ax2.set_title('Markers')
    ax3.imshow(labels, cmap='gray')
    ax3.axis('off')
    ax3.set_title('Segmentation')

    fig.tight_layout()
    plt.show()

def auto_segment(dt_array, threshold=1, visualize=False): #Should change to debug = False
    #Thresholding
    # The range of the binary image spans over (-1, 1).
    # We choose the hottest and the coldest pixels as markers.
    #print(dt_array)
    markers = np.zeros(dt_array.shape, dtype=np.uint)
    markers[dt_array <= threshold] = 1 #Must be dynamically chnange later
    markers[dt_array > threshold] = 2


    # Run random walker algorithm
    labels = random_walker(dt_array, markers, beta=10, mode='bf')

    #Show Graphic
    if(visualize):
        show_result(dt_array,markers,labels)

    return dt_array, markers, labels
