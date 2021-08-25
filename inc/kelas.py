import requests


class Shopee:

    def __init__(self):
        self.sess = requests.Session()

    def ambil_terlaris(self, key):
        url = 'https://shopee.co.id/api/v4/recommend/recommend?bundle=top_products_landing_page&intentionid=' + \
            key + '&limit=100&section=best_selling_sec'
        response = self.sess.get(url)
        data = response.json()
        return data
