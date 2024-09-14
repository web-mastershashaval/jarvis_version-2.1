import requests
import json



def get_news():
    url = 'pub_53390e661c1ea008d1c68d498187402928e24'
    news = requests.get(url).text
    news_dict = json.loads(news)
    articles = news_dict['articles']
    try:

        return articles
    except:
        return False


def getNewsUrl():
    return 'pub_53390e661c1ea008d1c68d498187402928e24'
