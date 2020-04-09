# Rapid-Labelling-Tool
A Python prototype tool for the rapid labeling of visual data to assist in the training of neural networks.

<img src="images/MLC-4.png" width="700">

Import your image data, define your annotations, and categorise your images. Output your resultant work in .json and/or .xlsx format.

<img src="images/MLC-2-2.png" width="700">

Define your number of categories and definitions per category, and then input as much (or little) information about each in each entry field.

Also includes second functionality mode, dubbed Object-location classification (OLC). You load your input imagery as usual, but this mode allows the user to define the objects of interest in a scene; the user is then taken to a graphical interface to draw bounding boxes around the elements present in an image. Outputs the top-left, bottom-right (x,y) co-ords of every image to .json and/or .xlsx format. 

<img src="images/OLC-Window_Defined.png" width="700">


Also featuring basic support for colourblind users: replaces the default red-green colour scheme with a colourblind-friendly one. 

<img src="images/Comparison_new.png" width="700">

Both a PIP environment ("env" folder) and Anaconda environment ("anaconda_env" folder) are included.

Note: The program only accepts input visual data (images) in .jpg file format. 
I've included some sample image data in "Sample_Images".

Submitted in partial fulfillment as part of BSc degree in Engineering with Management, TCD 2020
Daniel Danev
