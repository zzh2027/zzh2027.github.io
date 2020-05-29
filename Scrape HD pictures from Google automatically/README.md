## This class enables you to search for pictures by your key word and download automatically.
All pictures will be stored in the folder specified
You will required to download a Chrome driver first.

E.g.

from GoogleScraper import ScrapeGoogle

### folder1 has your Chrome.driver and folder2 is the folder where you would like to store your pictures
### If you are not aware of what Chrome driver is, go to this link:https://chromedriver.chromium.org/downloads
test = ScrapeGoogle(folder1, folder2)

### you wiil be required to pass the topic and number of pictures you want to download as parameters into this function and RUN
test.getPic(topic, num_pic = 50, pic_size = 'large')

### Enjoy!
### STAR me if this helps you at least a little.
