from bs4 import BeautifulSoup
import requests as req
from requests.exceptions import HTTPError


class HLTV:
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36'}
    URL = 'https://www.hltv.org'

    def __init__(self, URL='None'):
        self.URL = self.URL + URL  # Ссылка на карту игры
        self.statusUrl = self.status_url()  # Статус ссылке
        self.site = self.contents()  # Парсит страницы карты
        self.nameMaps = self.name_map()  # Название карты
        self.teamLeft = self.round_end_method_team1()  # Результат 1-й команды
        self.teamRight = self.round_end_method_team2()  # Результат 2-й команды
        self.overTime = self.overtime()  # Определяет был овертайм или нет
        self.methodWinRound = self.result() # Подсчет результата, без 1 и 16 раунда
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

    #: Парсинг страницы
    def contents(self):
        full_page = self.statusUrl
        try:
            soup = BeautifulSoup(full_page.content, 'lxml')

        except:
            return None
        else:
            return soup

    #: Название карты
    def name_map(self):
        try:
            map = self.site.find("div", class_="match-info-box-con").find("div", class_="match-info-box").find('div',
                                                                                                               class_='small-text').find_next().find_next().find_next().find_next().next_element.next_element.text

        except AttributeError:
            return 'MAP - NONE'

        except TypeError:
            return "TypeError - MAP"
        else:
            return map.split('\n')[1]

    # Результат по методам завершения раундов 1-й команды
    def round_end_method_team1(self):
        method_win_round = []
        try:
            round = self.site.find('div', class_='stats-section stats-match').find('div',
                                                                                   class_='standard-box round-history-con').find(
                'div', class_='round-history-team-row').find_all('img', class_='round-history-outcome')

            for item in round:
                src = str(item).split()[2]
                method = src.split('/')[-1].split('.')[0]
                method_win_round.append(method)

        except AttributeError:
            return 'TEAM 1 - NONE'

        except TypeError:
            return 'TypeError - TEAM 1'

        else:
            return method_win_round

    # Результат по методам завершения раундов 2-й команды
    def round_end_method_team2(self):
        method_win_round = []
        try:
            round = self.site.find('div', class_='stats-section stats-match').find('div',
                                                                                   class_='standard-box round-history-con').find(
                'div', class_='round-history-team-row').find_next_sibling().find_all('img',
                                                                                     class_='round-history-outcome')
            for item in round:
                src = str(item).split()[2]
                method = src.split('/')[-1].split('.')[0]
                method_win_round.append(method)

        except AttributeError:
            return 'TEAM 2 - NONE'

        except TypeError:
            return 'TypeError - TEAM 2'
        else:
            return method_win_round

    # Определяет был овертайм или нет
    def overtime(self):  # Проверка, есть овертайм или нет
        try:
            overtime = self.site.find('div', class_='stats-section stats-match').find_all(class_='standard-headline')
        except AttributeError as e:
            #print(f'Error(self.overTime): {e}')
            return None
        else:
            for item in overtime:
                if item.text == 'Overtime':

                    def team1_overtime():
                        static_over = self.site.find('div', class_='standard-box round-history-con').find_next(
                            class_='standard-box round-history-con').find(
                            'div', class_='round-history-team-row').find_all('img', class_='round-history-outcome')

                        ov_team_1 = []
                        for item in static_over:
                            src = str(item).split()[2]
                            method = src.split('/')[-1].split('.')[0]
                            ov_team_1.append(method)
                        return ov_team_1

                    def team2_overtime():
                        static_over = self.site.find('div', class_='standard-box round-history-con').find_next(
                            class_='standard-box round-history-con').find(
                            'div', class_='round-history-team-row').find_next_sibling().find_all('img',
                                                                                                 class_='round-history-outcome')

                        ov_team_2 = []
                        for item in static_over:
                            src = str(item).split()[2]
                            method = src.split('/')[-1].split('.')[0]
                            ov_team_2.append(method)
                        return ov_team_2

                    return team1_overtime(), team2_overtime()

            else:
                return None

    # Список методов побед в раундах
    def result(self):
        team_1 = []
        team_2 = []

        if self.teamLeft != 'TEAM 1 - NONE' and self.teamLeft != "TypeError - TEAM 1":
            team_1 = self.teamLeft

        if self.teamRight != 'TEAM 2 - NONE' and self.teamRight != 'TypeError - TEAM 2':
            team_2 = self.teamRight

        if self.overTime is not None:
            ov1, ov2 = self.overTime
            team_1 += ov1
            team_2 += ov2
        result_1 = team_1[1:15] + team_1[16:]
        result_2 = team_2[1:15] + team_2[16:]
        result_finall = result_1 + result_2
        return result_finall, result_1, result_2


#: ПРОВЕРКА
#hltv = HLTV('/stats/matches/mapstatsid/32415/liquid-vs-sk')
#hltv = HLTV('/stats/matches/mapstatsid/22280/fnatic-vs-envy')
#hltv = HLTV('/stats/matches/mapstatsid/145698/ence-academy-vs-prospects')
#hltv = HLTV('/stats/matches/mapstatsid/22268/ninjas-in-pyjamas-vs-virtuspro')
#print(hltv.statusUrl)
#print(hltv.site)
#print(hltv.nameMaps)
#print(hltv.teamLeft)
#print(hltv.teamRight)
#print(hltv.overTime)
#result_finall, result_1, result_2 = hltv.methodWinRound
#if result_finall != []:
    #print(result_finall)
    #print(result_1)
    #print(result_2)
