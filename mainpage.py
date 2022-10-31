from bs4 import BeautifulSoup
import requests as req
from requests.exceptions import HTTPError


class MatchCS:
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36'}

    def __init__(self, EVENTS='None'):
        self.event = EVENTS  # Номер события
        self.URL = f'https://www.hltv.org/results?event={EVENTS}'  # ЮРЛ Сайта
        self.statusUrl = self.status_url() # Статус сайта
        self.matchLink = self.links()  # Ссылки на матчи
        self.maps = self.map_links()  # Ссылки на карты в матчах
        self.countMaps = self.countmap()  # Подсчет кол-во карт
        self.eventName = self.even_name()  # Название чемпионата

    #: Проверяет статус ссылке
    def status_url(self):
        try:
            full_page = req.get(self.URL, headers=self.headers)
            full_page.raise_for_status()
        except HTTPError as http_err:
            return f'Произошла ошибка HTTP: {http_err}'  # Python 3.6
        except Exception as err:
            return f'Произошла другая ошибка: {err}'  # Python 3.6
        else:
            return full_page

    #: Парсит ссылке на матчи
    def links(self):
        full_page_list = self.statusUrl
        links = []
        try:
            soup = BeautifulSoup(full_page_list.content, 'lxml')
            match_page = soup.find('div', class_='results-holder').find('div', class_='results-all').find_all('a')

            for item in match_page:
                item_url = item.get('href')
                links.append(item_url)


        except AttributeError as e:
            print(f'Error(self.matchLink): {e}')
            return None
        else:
            return links

    #: Парсинг карт с матча
    def map_links(self):
        maps = self.matchLink
        map_cs = []
        if self.matchLink != None:
            for HLTV in maps:
                full_page_htlv = req.get('https://www.hltv.org' + HLTV, headers=self.headers)
                soup = BeautifulSoup(full_page_htlv.content, 'lxml')
                event = soup.find('div', class_='col-6 col-7-small').find('div', class_='flexbox-column').find_all('a')

                for item in event:
                    item_url = item.get('href')
                    map_cs.append(item_url)

            return map_cs
        else:
            return None

    #: Подсчет кол-во карт в чемпионате
    def countmap(self):
        if self.maps != None:
            count = self.maps
            return len(count)
        else:
            return 0

    #: Название чемпионата
    def even_name(self):
        try:
            full_page_list = req.get(self.URL, headers=self.headers)
            soup = BeautifulSoup(full_page_list.content, 'lxml')
            name_event = soup.find('div', class_='contentCol').find('div', class_='results').find('div',
                                                                                                  class_='event-hub').find(
                'div', class_='event-hub-title')

            return name_event.text

        except AttributeError:
            return 'EVENT-NONE'

        except TypeError:
            return "TypeError - EVENT"
        else:
            return name_event.text

# ПРОВЕРКА
#c = MatchCS(1666)
#print(c.statusUrl)
#print(c.matchLink)
#print(c.maps)
#print(c.eventName)
#print(c.countMaps)
