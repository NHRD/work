import urllib.request as req
from bs4 import BeautifulSoup as bs

def script_capture(url):

    company_names = []
    #description_title = []
    #descriptions = []

    #r = req.urlopen("https://fumasalse.com/search/?search_from_top=1&tab_btn=on&tab_btn_menu1=on&chu_code%5B%5D=28&chu_code%5B%5D=29&chu_code%5B%5D=31&tab_btn_data=on&listed=1&jugyoinsu%5B%5D=5&jugyoinsu%5B%5D=6")

    r = req.urlopen(url)
    soup = bs(r)
    com_name = soup.find_all(class_="s_res s_coprate")
    #des_title = soup.find_all(class_="searches__result__list__conts__text__heading")
    #des = soup.find_all(class_="searches__result__list__conts__text__excerpt")

    for i in range(len(com_name)):
        text_com = com_name[i].get_text()
        #text_dest = des_title[i].get_text()
        #text_des = des[i].get_text()
        company_names.append(str(text_com))
        #description_title.append(text_dest)
        #descriptions.append(text_des)

    r.close()
    
    #result = result.get_text()
    #print("ここから")
    #print(company_names[0])
    #print(description_title)
    #print(descriptions)
    #print("ここまで")

    return company_names #description_title, descriptions

#script_capture("https://fumasalse.com/search/?search_from_top=1&tab_btn=on&tab_btn_menu1=on&chu_code%5B%5D=28&chu_code%5B%5D=29&chu_code%5B%5D=31&tab_btn_data=on&listed=1&jugyoinsu%5B%5D=5&jugyoinsu%5B%5D=6")