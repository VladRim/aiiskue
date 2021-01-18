import requests

url = 'http://b2blk.megafon.ru'

# Важно. По умолчанию requests отправляет вот такой
# заголовок 'User-Agent': 'python-requests/2.22.0 ,  а это приводит к тому , что Nginx
# отправляет 404 ответ. Поэтому нам нужно сообщить серверу, что запрос идет от браузера

user_agent_val = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36'

# Создаем сессию и указываем ему наш user-agent
session = requests.Session()
r = session.get(url, headers = {
    'User-Agent': user_agent_val
})

# Указываем referer. Иногда , если не указать , то приводит к ошибкам.
session.headers.update({'Referer': url})

#Хотя , мы ранее указывали наш user-agent и запрос удачно прошел и вернул
# нам нужный ответ, но user-agent изменился на тот , который был
# по умолчанию. И поэтому мы обновляем его.
session.headers.update({'User-Agent': user_agent_val})

# Получаем значение _xsrf из cookies
_xsrf = session.cookies.get('_xsrf', domain=".b2blk.megafon.ru")

# Осуществляем вход с помощью метода POST с указанием необходимых данных
post_request = session.post(url, {
     'backUrl': 'https://b2blk.megafon.ru/subscribers/mobile',
     'username': 'CP_9250957589',
     'password': '115422822',
     '_xsrf':_xsrf,
     'remember':'yes',
})

#Вход успешно воспроизведен и мы сохраняем страницу в html файл
with open("lk_megafon.html","w",encoding="utf-8") as f:
    f.write(post_request.text)
#print(post_request.text)


url_list = 'https://b2blk.megafon.ru/subscriber/info/146222403'
r_number = session.get(url_list, headers = {
    'User-Agent': user_agent_val
})

#r = requests.get(url_list)
print(r.text)
