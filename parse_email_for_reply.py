from bs4 import BeautifulSoup
import pymssql as mc


class GetInputText():
    def __init__(self, html_from, input_text, namefolder, body):
        self.html_from = html_from
        self.input_text = input_text
        self.body = body
        self.id_cell = namefolder

    def parse_text_input(self):
        textFrom = self.html_from
        sup_txfr = BeautifulSoup(textFrom, 'html.parser')

        div_tag = sup_txfr.find("body")
        alex_tag = div_tag.find("p")
        new_tag = sup_txfr.new_tag("p")
        hr_tag = sup_txfr.new_tag("hr")
        new_tag.string = self.input_text
        alex_tag.insert_after(hr_tag)
        alex_tag.insert_after(new_tag)

        return sup_txfr.body

    def parse_body(self):
        soup = BeautifulSoup(self.body, 'html.parser')
        for index, img in enumerate(soup.findAll('img')):
            cid = img['src'][4:12] + '.png'
            img['src'] = f'http://10.1.0.31:9800/app_manager/attachments/{self.id_cell}/inline/{cid}'

        div_tag = soup.find('div')
        div_tag.insert_before(self.parse_text_input())
        my_html_string = str(soup).replace("'", '')

        return my_html_string

    def edit_body(self):
        mydb = mc.connect(server=SERVERMSSQL, user=USERMSSQL, password=PASSWORDMSSQL,
                          database=DATABASEMSSQL)

        mycursor = mydb.cursor()
        sql_select_query = mycursor.execute(
            f"""SELECT html_body FROM email WHERE id = {self.namefolder}""")
        result_ = mycursor.fetchall()
        body = self.parse_body(self.id_cell, result_[0][0])
        return body
