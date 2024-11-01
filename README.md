# drexel-course-web-scraper
Fetches Drexel University's public class information and returns a CSV file.
Example of scrapeable webpage: https://termmasterschedule.drexel.edu/webtms_du/collegesSubjects/202415?collCode=

The program is just one python file that goes to as specified term and gets all schedules for every class that is synchronous (i.e. the time listed on the term master schedule is not "TBD")

To get past the authentication, you need to be a Drexel student with a Drexel account and have access to the page linked above. 
Create a .env file in the same directory as main.py with these variables:\n
DREXEL_USERNAME=YOUREMAIL\n
DREXEL_PASSWORD=YOURPASSWORD\n

Fill it in with your credentials and when running wait for it to enter your password then press enter or the sign in button.
It will then wait, for up to 5 minutes, for you to authenticate through whatever MFA method you have set up.
Then just let it run until completion.
