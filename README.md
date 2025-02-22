# ISF Watchkeeper Report Reader
The script `extract_watchkeeper.py` reads reports from ISF Watchkeeper software such as the one shown below, calculates some statistical values such as mean by positions and monthly periods, and prepares Excel reports with these values. This is done with the help of image processing techniques, OCR, and regular expressions.

![sample page](sample.jpg)

The reports are assumed to be of the same ship. Each function in the script is thoroughly commented with all the input and outputs.

# Instructions
Clone the repository and install the necessary libraries using `pip install -r requirements.txt`. Then, run the main script `extract_watchkeeper.py`, and use the function of your choice.
