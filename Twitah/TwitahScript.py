import tweepy
import datetime
from tweepy import OAuthHandler
from xlwt import Workbook
import RAKE
import sys
import time


# Variables de acceso
access_token = ""
access_token_secret = ""
consumer_key = ""
consumer_secret = ""


def leerTimeLine():
    listaTweets = []
    for elemento in limit_handled(tweepy.Cursor(api.user_timeline, screen_name = "Razorbacks_UBU", include_rts = True).items(1000)):
        listaTweets.append(elemento)

    return listaTweets

def limit_handled(cursor):
    while True:
        try:
            print("Devuelve")
            yield cursor.next()
        except StopIteration:
            return True
        except:
            a = sys.exc_info()
            print(a)
            time.sleep(16 * 60)


def creaExcel(lista):
    fila = 1
    wb = Workbook()

    hoja1 = wb.add_sheet(datetime.date.today().__str__())

    hoja1.write(0, 0, "Fecha")
    hoja1.write(0, 1, "Texto")
    hoja1.write(0, 2, "Enlace")
    hoja1.write(0, 3, "RT")
    hoja1.write(0, 4, "Respuestas")
    hoja1.write(0, 5, "Favs")
    hoja1.write(0, 6, "Tema")
    reversed(lista)
    for elemento in lista:
        hoja1.write(fila, 0, elemento.created_at.__str__())
        hoja1.write(fila, 1, elemento.text)
        hoja1.write(fila, 2, 'https://twitter.com/' + elemento.user.name + '/status/' + elemento.id_str)
        hoja1.write(fila, 3, elemento.retweet_count)
        #hoja1.write(fila, 4, elemento.reply_count)

        hoja1.write(fila, 5, elemento.favorite_count)
        hoja1.write(fila, 6, keywordExtraction(elemento.text))
        #hoja1.write(fila, 4, indiceMagicoDePol())
        fila = fila + 1
    wb.save(
        'excelTwitterOGSeries' + datetime.date.today().__str__() + ' ' + datetime.datetime.today().hour.__str__() + '-' + datetime.datetime.today().minute.__str__() + '-' + datetime.datetime.today().second.__str__() + '.xls')


def keywordExtraction(texto):
    r = RAKE.Rake("SpanishStoplist.txt")
    lista = r.run(texto, maxWords=3)
    return lista.__str__()


# def indiceMagicoDePol():p


if __name__ == '__main__':
    try:
        # This handles Twitter authetification and the connection to Twitter Streaming API
        auth = OAuthHandler(consumer_key, consumer_secret)
        auth.set_access_token(access_token, access_token_secret)
        api = tweepy.API(auth)

        lista = leerTimeLine()
        creaExcel(lista)
    except:
        e = sys.exc_info()
        print(e)

