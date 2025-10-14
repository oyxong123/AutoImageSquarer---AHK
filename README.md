### _**This repo is obsolete and no longer maintained.**_

# Auto Image Squarer
<p align="center">
    <img width="683" height="363" alt="image" src="https://github.com/user-attachments/assets/2d2ee59e-d49a-4dfb-bf42-1c4b55c73bb0" />
</p>

Auto Image Squarer is a utility program that allows me to quickly grab a folder and convert all of the images within it into 1:1 ratio without sacrificing quality and completeness of each of the original images. 

It utilizes an AHK script that accepts a specific hotkey to start the program, select a folder, and trigger the python script. 
The python script will then interact with an image squaring website through Selenium WebDriver to convert the batch of images into 1:1 ratio images. 
Lastly, the program will proceed to download the 1:1 ratio images and save them inside the folder with the original images. 

I personally used it to square 9000+ images with it. 
It, along with the other programs and applications I have built, integrates into one large ecosystem. 

Feel to to try it out, share it with others, and let me know if this program personally helped you in any way!

Contributions, issues, and pull requests are welcomed!

## Requirements
- Windows 10 (other versions not tested)
- [AutoHotKey v1.1](https://www.autohotkey.com/)
- Python 3.6 or above
- Suitable Chromedriver version based on your current Google Chrome version
- Python packages `selenium`, `win10toast-persist`, and `webdriver-auto-update`

