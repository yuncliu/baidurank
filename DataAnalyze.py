#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from Rank import Rank
import xlsxwriter
import csv

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
        self.websitesRankXLSX.append(t)
        for site in websites:
            a = [site]
            for kw in keywords:
                rank = 20
                for i in range(len(self.rawRank[kw])):
                    if self.rawRank[kw][i].find(site) >= 0:
                        rank = i + 1
                a.append(rank)
            a.append(sum(a[1:])/len(a[1:]))#calculage average rank
            self.websitesRankXLSX.append(a)

    def dumpWebstitesStasticsToXLSX(self, filename):
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        for row in range(len(self.websitesRankXLSX)):
            r = self.websitesRankXLSX[row]
            for col in range(len(r)):
                worksheet.write(row, col, r[col])

        workbook.close()

def main():
    a = RankData()
    a.getRanks()
    a.dumpToFile('r.csv')
    a.dumpWebstitesStasticsToCSV('w.csv')
    a.dumpWebstitesStasticsToXLSX('w.xlsx')

if __name__ == "__main__":
    main()
