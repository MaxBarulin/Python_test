from bs4 import BeautifulSoup
from requests_html import HTMLSession

def main(url):
    chrome_log = 'barulin.ma:Maxim352678999@'
    src_segment_1 = url[:7]
    src_segment_2 = url[7:]
    SRC = f'{src_segment_1}{chrome_log}{src_segment_2}'
    
    session = HTMLSession()

    r = session.get(SRC)

    r.html.render()
    titles = r.html.find('.fieldValue')
    titles_name = r.html.find('.name')
    titles_src = r.html.find('iframe')
    try:
        titles_src_1 = str(titles_src[0])
        titles_src_2 = titles_src_1.split("'")
        r1 = session.get(titles_src_2[3])
        iframe = str(r1.content)
        iframe_1 = iframe.replace('/', '').replace('<', '').replace('>', '').replace(' ', '').split('td')
        iframe_2 = iframe_1[1].split('=')
        try:
            color = iframe_2[2]
        except:
            color = iframe_2[1]
    except:
        color = 'orange'
        
    if color == 'MediumSeaGreen':
        color = 'НА ТЕРРИТОРИИ'
    elif color == 'tomato':
        color = 'ОТСУТСТВУЕТ'
    elif color == 'orange':
        color = 'НЕ РАБОТАЕТ'

    print(color)
        
    name = titles_name[0].text
    management = titles[0].text
    departament = titles[1].text
    post = titles[2].text
    eMail = titles[3].text
    phone = titles[4].text
    location = titles[6].text
    depNum = titles[8].text
    idNum = titles[9].text
    print(f'имя: {name}({idNum})\nподразделение: {management}({depNum})\nотдел: {departament}\nдолжность: {post}\nлокация: {location}\nтелефон: {phone}\nпочта: {eMail}')

    r.close
    #for title_name in titles_name:
    #    print(title_name.text)
        
    ##for title in titles:
    ##    print(title.text)

if __name__ == '__main__':
    src_1 = input('вставь ссылку: ')
    main(src_1)






    
