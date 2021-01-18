from bs4 import BeautifulSoup
import lxml

with open("31408497.html", "r") as f:
    contents = f.read()

    soup = BeautifulSoup(contents, 'lxml')

print(soup.h2)

