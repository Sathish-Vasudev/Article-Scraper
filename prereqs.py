import subprocess
import sys

print("This pre-requisites program will install all the necessary modules.\n")

def install(name):
    subprocess.call([sys.executable, '-m', 'pip', 'install', name])
    
install("newspaper3k")
install("python-docx")
import nltk
nltk.download()

print("\nAll the modules are installed !")
print("You can now run the Article Scraper program!")

