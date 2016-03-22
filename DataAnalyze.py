#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from Rank import Rank
import xlsxwriter
import csv
import datetime

keywords = ['nokia', 'python']
websites =['nokia', 'ifanr', '2cto']

class RankData(object):
    """Docstring for RankData. """
    def __init__(self):
        """TODO: to be defined1. """
        self.rawRank = dict()
        self.websitesRankCSV = dict()
        self.websitesRankXLSX = []

    def getRanks(self):
        for i in keywords:
            x = Rank(i)
            self.rawRank[i] = x.getRank()

        self.countWebsitesCSV()
        self.countWebsitesXLSX()

    def dumpToFile(self, filename):
        a = list()
        [a.append(i) for i in self.rawRank.values()]
        csvrows = list(zip(*a))
        with open(filename, 'w') as csvfile:
            fieldnames = keywords
            writer = csv.DictWriter(csvfile, fieldnames = fieldnames)
            writer.writeheader()
            for i in csvrows:
                t = dict(zip(fieldnames, i))
                writer.writerow(t)

    def countWebsitesCSV(self):
        for site in websites:
            a={'Name': site}
            for kw in self.rawRank.keys():
                rank = 20
                for i in range(len(self.rawRank[kw])):
                    if self.rawRank[kw][i].find(site) >= 0:
                        rank = i + 1
                a[kw] = rank
            self.websitesRankCSV[site] = a

    def dumpWebstitesStasticsToCSV(self, filename):
        fieldnames = ['Name']
        [fieldnames.append(i) for i in keywords]
        [print(i) for i in fieldnames]
        with open(filename, 'w') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames = fieldnames)
            writer.writeheader()
            for v in self.websitesRankCSV.values():
                writer.writerow(v)

    def countWebsitesXLSX(self):

        """keywords title"""
        t = ['']
        [t.append(i) for i in keywords]
        t.append('average')
        t.append('rank')
        self.websitesRankXLSX.append(t)
        average = []
        for site in websites:
            a = [site]
            for kw in keywords:
                rank = 20
                for i in range(len(self.rawRank[kw])):
                    if self.rawRank[kw][i].find(site) >= 0:
                        rank = i + 1
                a.append(rank)
            a.append(sum(a[1:])/len(a[1:]))#calculage average rank
            average.append(sum(a[1:])/len(a[1:]))
            self.websitesRankXLSX.append(a)
        average.sort()
        averagedict = {}
        for i in range(len(average)):
            averagedict[average[i]] = i + 1

        for l in self.websitesRankXLSX:
            last = l[len(l)-1]
            if isinstance(last, str):
                continue
            l.append(averagedict[last])


    def dumpWebstitesStasticsToXLSX(self, filename):
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()

        title_format = workbook.add_format({
            'bold':     True,
            'align':    'center',
            'valign':   'vcenter',
            'font_size': 20,
        })

        second_row_format = workbook.add_format({
            'bold':     False,
            'align':    'left',
            'valign':   'vbottom',
        })
        x_format = workbook.add_format({
            'bold':     False,
            'border':   1,
            'diag_type':   2,
            'align':    'center',
            'align':   'vcenter',
            'text_wrap':True,
        })
        c_format = workbook.add_format()
        c_format.set_align('center')
        c_format.set_align('vcenter')
        c_format.set_border(1)

        worksheet.merge_range('A1:U1', '百度快照排名统计表', title_format)
        worksheet.set_row(0, 30)
        worksheet.merge_range('A2:U2', datetime.datetime.now().strftime('统计日期：%Y年%m月%d日%H:%M'),
                second_row_format)
        worksheet.set_row(1, 20)
        worksheet.set_row(2, 20)
        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:B', 2)

        worksheet.merge_range('A3:A4', '      关键词\n网站   ', x_format)
        worksheet.write('B3', '序', c_format)
        for i in range(len(keywords)):
            worksheet.write(2, 2 + i , (i + 1).__str__(), c_format)

        worksheet.write('B4', '词')

        for row in range(len(self.websitesRankXLSX)):
            r = self.websitesRankXLSX[row]
            for col in range(len(r)):
                if row == 0 and col == 0:
                    worksheet.write(row + 3, col, r[col], x_format)
                    worksheet.set_row(row + 3, 30)
                else:
                    worksheet.set_row(row + 3, 20)
                    if col == 0:
                        #worksheet.write(row + 3, col + 1, r[col])
                        worksheet.merge_range(row + 3, col, row + 3, col + 1, r[col], c_format)
                    else:
                        worksheet.write(row + 3, col + 1, r[col], c_format)

                    if r[col] =='average' or r[col] == 'rank':
                        worksheet.merge_range(row + 3, col + 1, row + 2, col + 1, r[col], c_format)

        last_row = len(self.websitesRankXLSX) + 5
        worksheet.merge_range('A{0}:U{0}'.format(last_row), '注：此数据统计出现在百度快照头两页的排名按实际排名计入，出现在第三页之后的排名均按排名20位计入。')

        workbook.close()

def main():
    a = RankData()
    a.getRanks()
    a.dumpToFile('r.csv')
    a.dumpWebstitesStasticsToCSV('w.csv')
    a.dumpWebstitesStasticsToXLSX('w.xlsx')

if __name__ == "__main__":
    main()
