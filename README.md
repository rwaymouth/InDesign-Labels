InDesign-Labels v1.2
===============

Script to label every linked image in a set of InDesign files with their filename, and then output a PDF.


### Updates:
- Now color codes labels by the location of the image. If the linked image is in any folder other than FPOs or Downloads it will be labeled in red. 
- The script now appends the location of the non-FPO/Download images to the front of the label, if it's in any of the standard folders (CD, ADS, CLIENT, STANDING ART, etc.)
- The script now does a better job of labeling the images without overlapping other labels. It's still not perfect, but it's significantly better.
### Known issues:
- Occasionally the script will run into a missing font issue the first time you run it. 9 times out of 10 if you run the script again it will work. So if the script fails the first time try it again before giving up, it will probably work fine. 
