from requests import Session
from bs4 import BeautifulSoup as sp
from threading import Thread
from time import sleep
from os import system


headers: dict = {
    'sec-ch-ua-platform': '"Linux"',
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
    'sec-ch-ua-mobile': '?0',
}

class Goaloo:
    net: Session = Session()
    net.headers = headers
    
    # constructor
    def __init__(self, id:str):
        self.url: str = 'https://www.goaloo14.com/'
        self.id = id
        self.api_odds: str = 'ajax/soccerajax'
        self.path_details: str = 'match/live-{}'
        self.path_analysis: str = 'match/h2h-{}'
        self.datas: dict = {}
        self.home_results: list = []
        self.away_results: list = []
        
    # this function is for get the odds in Bet365
    def get_odds(self) -> None:
        params: dict = {
            'type': '14',
            't': '1',
            'id': f'{self.id}',
        }
        response: dict = self.net.get(self.url+self.api_odds, params=params).json()
        self.datas['id'] = f'{self.id}'
        self.datas['early'] = response["Data"]["mixodds"][0]["euro"]["f"]
        
    # this function is for get all of goals
    def get_goals(self) -> None:
        list_goals: list = []
        response: object = self.net.get(self.url+self.path_details.format(self.id))
        soup: sp = sp(response.text, 'html.parser')
        goals: list = soup.select('tr[align="center"]')
        for goal in goals:
            if goal.select_one('span[class="ky_score"]'):
                list_goals.append(goal.select_one('b').text)
        list_goals = list_goals[::-1]
        self.datas['goals'] = list_goals
        
    # this function is for get the date and the time about game and league and team name
    def get_time_league(self) -> None:
        response: Session = self.net.get(self.url+self.path_details.format(self.id))
        soup: sp = sp(response.text, 'html.parser')
        time: str = soup.select_one('span[name="timeData"]').get('data-t').split(' ', 1)
        league: str = soup.select_one('span[class="LName"]').text
        team: list = soup.select('div[class="sclassName"]')
        self.datas['time'] = {'hour': time[1], 'date': time[0].replace('/', '-')}
        self.datas['league'] = league.replace('\n', '')
        self.datas['team'] = f'{team[0].text} vs {team[1].text}'
        
        
    # this function is for get the last 5 games
    def get_games(self, list_data: list, table:str) -> None:
        response: object = self.net.get(self.url+self.path_analysis.format(self.id))
        soup: sp = sp(response.text, 'html.parser')
        results1: list = soup.select('table[id="{}"] td[onclick*="soccerInPage.detail"] span[class*="fscore"]'.format(table))[0:5]
        results2: list = soup.select('table[id="{}"] td[onclick*="soccerInPage.detail"] span[class*="hscore"]'.format(table))[0:5]
        situation: list = soup.select('table[id="{}"] td[class="hbg-td1"] span'.format(table))[0:5]
        for position in range(len(situation)):
            list_data.append(f'{results1[position].text}-{results2[position].text}{situation[position].text}')
    
    # this function is for set the last 5 games in the dict
    def set_games(self) -> None:
        Thread(target=self.get_games, args=(self.home_results, 'table_v1',)).start()
        Thread(target=self.get_games, args=(self.away_results, 'table_v2',)).start()
        sleep(5)
        self.datas['results'] = {'Home': self.home_results, 'Away': self.away_results}
        #print(self.datas)
    
    def start(self) -> None:
        Thread(target=self.get_time_league).start()
        Thread(target=self.get_goals).start()
        Thread(target=self.get_odds).start()
        self.set_games()
        sleep(15)
        try:
            system('clear')
        except:
            system('cls')
        return self.datas