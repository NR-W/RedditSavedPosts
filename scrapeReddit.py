#import selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
# Writing to an excel sheet using Python
import xlwt
from xlwt import Workbook
# time delay
import time
import enum


### Constants and configuration

# Reddit username, as it appears in the URL for the page listing your saved posts.
reddit_username = "<username>"

COL_IDX_POST_TITLE=0
COL_IDX_POST_SUBREDDIT=1
COL_IDX_POST_LINK=2
COL_IDX_POST_EXTERNAL_LINK=3    # if applicable

# creating enumerations (using class) for the different types of data that we want to import.
class CellDataType(enum.Enum):
    text = 1
    href_hyperlink = 2
    external_hyperlink = 3

### Functions

# For a particular saved post (savedPostsListp[index]), get a particular piece of data (e.g. post title),
# which is pointed to by the relative XPath (dataXPath), and add it to the spreadsheet
# in the corresponding column (excelCol), formatted as required for the
# specific type of data we are importing (cellContentType)
def setPostDataToExcel(savedPostsList, index, dataXPath, excelCol, cellContentType):
    postTitles = savedPostsList[index].find_elements(by=By.XPATH, value=dataXPath)
    titleCount = len(postTitles)
    for idx2 in range(titleCount):
        if cellContentType == CellDataType.href_hyperlink:
            cellContent = "HYPERLINK(\""+postTitles[idx2].get_attribute('href')+"\", \"POST LINK\")"
        elif cellContentType == CellDataType.text:
            cellContent = postTitles[idx2].text
        elif cellContentType == CellDataType.external_hyperlink:
            cellContent = "HYPERLINK(\""+postTitles[idx2].text+"\", \"EXTERNAL LINK\")"

        if cellContentType == CellDataType.text and excelCol == COL_IDX_POST_TITLE:
            sheet1.write(index+1, excelCol, cellContent, titleColStyle)
        elif cellContentType == CellDataType.text:
            sheet1.write(index+1, excelCol, cellContent)
        elif cellContentType == CellDataType.href_hyperlink or cellContentType == CellDataType.external_hyperlink:
            sheet1.write(index+1, excelCol, xlwt.Formula(cellContent))
        break

# Import a hyperlink. Note: the dataXPath should be pointing to a <a> tag.
def setPostLink(savedPostsList, index, dataXPath, excelCol):
    setPostDataToExcel(savedPostsList, index, dataXPath, excelCol, CellDataType.href_hyperlink)

# Import an external link to which the post is pointing.
# This function is deprecated. This is because the external links (due to formatting in Reddit) will need to be imported using the setPostLink() function.
def setPostExternalLink(savedPostsList, index, dataXPath, excelCol):
    setPostDataToExcel(savedPostsList, index, dataXPath, excelCol, CellDataType.external_hyperlink)

# Import text from the saved post.
def setPostTextToExcel(savedPostsList, index, dataXPath, excelCol):
    setPostDataToExcel(savedPostsList, index, dataXPath, excelCol, CellDataType.text)

# Import the title of the saved post.
def setPostTitle(savedPostsList, index):
    setPostTextToExcel(savedPostsList, index, titleXPath, COL_IDX_POST_TITLE)


processedPosts = 0  #how many posts we have processed so far at this point
def importRedditSaves():
    global processedPosts
    savedPosts = driver.find_elements(by=By.XPATH, value=itemXPath)
    postCount = len(savedPosts)
    print("postCount: ", postCount)

    for idx in range(processedPosts, postCount):
        processedPosts += 1
        setPostLink(savedPosts, idx, externalLinkXPath, COL_IDX_POST_EXTERNAL_LINK)
        setPostLink(savedPosts, idx, linkXPath, COL_IDX_POST_LINK)
        setPostTextToExcel(savedPosts, idx, subRedditXPath, COL_IDX_POST_SUBREDDIT)
        setPostTextToExcel(savedPosts, idx, titleXPath, COL_IDX_POST_TITLE)


# Initialize Output
wb = Workbook()
sheet1 = wb.add_sheet('Posts')
# Column titles/headers
sheet1.write(0, COL_IDX_POST_TITLE, 'Post Titles')
sheet1.write(0, COL_IDX_POST_SUBREDDIT, 'Post Sub-Reddit')
sheet1.write(0, COL_IDX_POST_LINK, 'Post Link')
sheet1.write(0, COL_IDX_POST_EXTERNAL_LINK, 'Post External Link')

# Format the Excel Sheet
titleColStyle = xlwt.XFStyle()
titleColStyle.alignment.wrap = titleColStyle.alignment.WRAP_AT_RIGHT

sheet1.col(0).width = 256*75
sheet1.col(1).width = 256*25
sheet1.col(2).width = 256*15
sheet1.col(3).width = 256*25


# Scraper
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get('https://www.reddit.com/user/'+reddit_username+'/')
username = input("Enter to continue... ") # Wait for user input before continuing with scraping
driver.get('https://www.reddit.com/user/'+reddit_username+'/saved/')

itemXPath = "/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[3]/div[1]/div[2]/div[1]/div"

# Relative XPaths
titleXPath = "./div/div/div[2]/div/div[2]/div[1]/div[1]/a/div/h3"
linkXPath = "./div/div/div[2]/div/div[2]/div[1]/div[1]/a"
subRedditXPath = "./div/div/div[2]/div/div[2]/div[2]/div[1]/a"
externalLinkXPath = "./div/div/div[2]/div/div[2]/div[1]/a"

# Keep scrolling down (and loading saved posts) until we have
# reached the end of the list of posts and there are no more posts to load (all saved posts are loaded).
xPathMoreItemsCard = "/html/body/div[1]/div/div[2]/div[2]/div/div/div/div[2]/div[3]/div[1]/div[2]/div[2]"
moreItemsElem = driver.find_elements(by=By.XPATH, value=xPathMoreItemsCard)
moreItemsElemCount = len(moreItemsElem)
print("Searching for last saved item...", end="")
try:
    while moreItemsElemCount > 0:
        print(".", end="")
        importRedditSaves() # import visible posts before loading/scrolling the next set
                            # if we don't do it this way, many posts do not import correctly for some reason.

        driver.find_element(by=By.TAG_NAME, value='body').send_keys(Keys.END)
        moreItemsElem = driver.find_elements(by=By.XPATH, value=xPathMoreItemsCard)
        moreItemsElemCount = len(moreItemsElem)
        time.sleep(3)
except Exception as e:
    print(repr(e))

print("Reached end of list...")
driver.find_element(by=By.TAG_NAME, value='body').send_keys(Keys.HOME)
time.sleep(3)

driver.quit()   # close the webdriver

# Save and close the Excel file
wb.save('Reddit Saved Posts.xls')
