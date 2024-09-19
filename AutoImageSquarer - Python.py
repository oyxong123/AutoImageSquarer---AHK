# This script is only the Python part (2/2) of the full script, the other half is at AutoImageSquarer - AHK.ahk
import time
import datetime
import os
import zipfile
import shutil
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_auto_update.webdriver_auto_update import WebdriverAutoUpdate
from win10toast_persist import ToastNotifier

# Wait for files to finish download in specified path.
def WaitFileDownloaded(path: str):
    fileNames = []
    filesWithKeywords = []
    keywordNotDownloaded = [".crdownload", ".download"]
    initialBufferTime = 3
    subsequentBufferTime = 1
    loopCount = 0
    time.sleep(initialBufferTime)   
    while len(fileNames) == 0 or len(filesWithKeywords) != 0:
        fileNames = os.listdir(path)
        filesWithKeywords = [y for x in keywordNotDownloaded for y in fileNames if x in y]  # Checks if any of the file names contain the keywords as substring and make them into a list.
        time.sleep(subsequentBufferTime)
        loopCount += 1
        if loopCount >= 50:  
            # Incidentally, square image website has a timeout of 30 secs and will show error if not finish uploading within that timeframe.
            # Need to account for the downloading time as well: Added 25 secs.
            raise Exception("File error by website occured.")
        
def GetScriptArgument(index):
    try:
        sys.argv[index]
    except IndexError:
        return None
    else:
        return sys.argv[index]
    
loggedInUser = os.getlogin()
pathDownload = f"C:/Users/{loggedInUser}/Downloads/squared-images"
pathDownloadAlt = f"C:\\Users\\{loggedInUser}\\Downloads\\squared-images"
pathProgrammingLife = f"C:/Users/{loggedInUser}/Desktop/Programming Life"
pathLanguageCottage = f"C:/Users/{loggedInUser}/Desktop/Language Cottage"

try:
    # Set script initialize state.
    scriptStatus = True
    
    # Get and validate arguments passed from the AHK script.
    squaringFolder = GetScriptArgument(1)
    pathArgument = GetScriptArgument(2)
    if squaringFolder not in ("p", "l", "manual"):
        raise Exception("Invalid argument.")
    if squaringFolder == "" and pathArgument == None:
        raise Exception("Manually selected path not provided.")
    
    # Determine squaring images.
    strDateToday = datetime.date.today().strftime("%Y_%m_%d")
    if squaringFolder == "p":  # Programming life
        pathCategory = pathProgrammingLife
        pathToday = os.path.join(pathCategory, strDateToday)
    elif squaringFolder == "l":  # Language cottage
        pathCategory = pathLanguageCottage
        pathToday = os.path.join(pathCategory, strDateToday)
    else:
        pathToday = pathArgument  # Manually selected folder
    pathSquaredToday = os.path.join(pathToday, "squared")
    
    # Ensure all important paths exist and are empty.
    paths = [pathDownload, pathSquaredToday]
    for path in paths:
        if not os.path.exists(path):
            os.makedirs(path)
        else:
            shutil.rmtree(path, ignore_errors=True)
            os.makedirs(path)
      
    # Notify computer squaring has started for what category.
    notif = ToastNotifier()
    if squaringFolder == "manual":
        notif.show_toast("Auto Image Squarer", f"Squaring started.\nImage squared: -\nCategory: Manual", duration=5)
    else:
        notif.show_toast("Auto Image Squarer", f"Squaring started.\nImage squared: -\nCategory: {os.path.basename(pathCategory)}", duration=5)
        
    # Autoupdate chromedriver.
    driverDirectory = "C:/Windows"
    # WebdriverAutoUpdate(driverDirectory).check_driver()  # removed due to not working correctly.
    
    # Configure driver options and services.
    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'normal'
    options.add_argument("--headless")
    options.add_argument("--window-size=1920x1080")
    prefs = {"download.default_directory" : pathDownloadAlt}  # Set squared-images zip files download path.
    options.add_experimental_option("prefs",prefs)
    logPath = "C:/Users/Admin/Documents/Selenium/AutoImageSquarer"
    service = webdriver.ChromeService(service_args=['--log-level=DEBUG', '--append-log', '--readable-timestamp'], log_output=logPath)
    driver = webdriver.Chrome(options, service)
    driver.implicitly_wait(20)
    
    # Access image squarer website.
    driver.get("https://squaremyimage.com/")

    # Get paths of squaring images.
    pathFiles = []
    dirPathToday = os.listdir(pathToday)
    for fileName in dirPathToday:
        pathFile = os.path.join(pathToday, fileName)
        if os.path.isfile(pathFile):
            pathFiles.append(pathFile)
            
    # Check if selected folder contains no files.
    if len(pathFiles) <= 0:
        errorMessage = f'"{pathToday}" has no files.'
        raise Exception(errorMessage)
    
    # Submit squaring images.
    for pathFile in pathFiles:
        websiteFileInput = driver.find_element(By.CSS_SELECTOR, "input[type=file]")
        websiteFileInput.send_keys(pathFile)
    button = driver.find_element(By.ID, "submit")
    driver.execute_script('arguments[0].click();', button)

    # Wait for the file to finish download.
    WaitFileDownloaded(pathDownload)

    # Determine the latest downloaded squared-images zipfile.
    [latestFileDownloaded] = os.listdir(pathDownload)
            
    # Extract and delete the latest downloaded squared-images zipfile.
    pathFileDownload = os.path.join(pathDownload, latestFileDownloaded)
    with zipfile.ZipFile(pathFileDownload, 'r') as zippedFile:
        shutil.rmtree(pathSquaredToday, ignore_errors = True)
        zippedFile.extractall(pathSquaredToday)
    shutil.rmtree(pathDownload, ignore_errors = True)
    
    # Check if squared file count matches with original.
    lenFileSquared = len(os.listdir(pathSquaredToday))
    lenFileOriginal = len(pathFiles)
    if lenFileSquared != lenFileOriginal:
        raise Exception("Not all files are squared!")
    
except Exception as exception:
    # Catch exception.
    scriptStatus = False
    errorMessage = str(exception)
    
    # Remove partially-created folders.
    if os.path.exists(pathDownload):
        shutil.rmtree(pathDownload, ignore_errors = True)
        
# Add a notification to notify me about the success or failure.
notif = ToastNotifier()
if scriptStatus:
    if squaringFolder == "manual":
        notif.show_toast("Auto Image Squarer", f"Successfully squared all images.\nImage squared: {lenFileSquared}\nCategory: Manual", duration=5)
    else:
        notif.show_toast("Auto Image Squarer", f"Successfully squared all images.\nImage squared: {lenFileSquared}\nCategory: {os.path.basename(pathCategory)}", duration=5)
    print("Script completed. Operation Success. All image squared.") 
else:
    notif.show_toast("Auto Image Squarer", "Operation failed.\n" + errorMessage, duration=5)
    print("Script failed. Operation failed.\n" + errorMessage) 

