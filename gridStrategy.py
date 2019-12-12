# coding: utf8

import xlwt
# 网格策略

class GridCell:

    # 档位
    rank = None

    # 买入金额
    buyAmount = None

    # 卖出金额
    saleAmount = None

    # 买入数量
    buyNum = None

    # 卖出数量
    saleNum = None

    # 买入价
    buyPrice = None

    # 卖出价
    salePrice = None

    # 类型
    type = None

    # 利润
    profit = None

    # 利润百分比
    profitRate = None

    # 手续费
    serviceCharge = 5

    def caculateProfit(self):
        self.serviceCharge=round(self.buyAmount * 0.00025)
        if self.serviceCharge < 5:
            self.serviceCharge = 5
        self.profit = round(self.saleAmount - self.buyAmount-self.serviceCharge)
        self.profitRate = round(self.profit / self.buyAmount, 2)

    def desc(self):
        print self.type, self.rank, self.buyPrice, self.buyNum, self.buyAmount, self.salePrice, self.saleNum, self.saleAmount, self.profit, self.profitRate


class MakeGrid:

    def __init__(self, dir):
        # 目标最高价
        self.max = 0.80

        # 大网
        self.large = 0.3

        # 中网
        self.middle = 0.15

        # 小网
        self.small = 0.05

        # 压力测试价
        self.pressPrice = 0

        # 每一小格递增买入比率
        self.stepRate = 0.15

        # 从哪一个档位开始递增投资
        self.initalStep = 2

        self.min = 0.4 * self.max

        # 初始买入数量
        self.minBuyNum = 1000

        self.smallGrids = []
        self.middleGrids = []
        self.largeGrids = []

        self.targetDir = dir+'/Grid.xls'


    def makeGrids(self):

        self.largeGrids = []
        self.middleGrids = []
        self.smallGrids = []

        # 档位对应买入数量
        rateNum = {}

        # 小网织网

        thisRank = 1.0
        index = 1
        thisBuyNum = self.minBuyNum
        while thisRank >= 0.4:
            cell = GridCell()
            cell.rank = thisRank
            cell.type = "小网"
            if index > self.initalStep:
                thisBuyNum = round((1 + self.stepRate) * thisBuyNum)
            cell.buyPrice = round(self.max * thisRank, 2)
            cell.buyNum = int(thisBuyNum)
            cell.buyAmount = round(cell.buyNum * self.max * thisRank)
            cell.salePrice = round(self.max * (thisRank + self.small), 2)
            cell.saleNum = int(thisBuyNum)
            cell.saleAmount = round((thisRank + self.small) * self.max * cell.saleNum)
            cell.caculateProfit()
            self.smallGrids.append(cell)
            index += 1
            thisRank = round(thisRank - self.small, 2)
            rateNum[thisRank] = thisBuyNum
            print thisRank

        print "rateNum", rateNum

        # 中网织网
        thisRank = round(1.0 - self.middle, 2)
        while thisRank >= 0.4:
            thisBuyNum = rateNum[thisRank]
            cell = GridCell()
            cell.rank = thisRank
            cell.type = "中网"
            cell.buyPrice = round(self.max * thisRank, 2)
            cell.buyNum = int(thisBuyNum)
            cell.buyAmount = round(cell.buyNum * self.max * thisRank)
            cell.salePrice = round(self.max * (thisRank + self.middle), 2)
            cell.saleNum = int(thisBuyNum)
            cell.saleAmount = round((thisRank + self.middle) * self.max * cell.saleNum)
            cell.caculateProfit()
            self.middleGrids.append(cell)
            thisRank = round(thisRank - self.middle, 2)

        # 大网织网
        thisRank = round(1.0 - self.large, 2)
        while thisRank >= 0.4:
            thisBuyNum = rateNum[thisRank]
            cell = GridCell()
            cell.rank = thisRank
            cell.type = "大网"
            cell.buyPrice = round(self.max * thisRank, 2)
            cell.buyNum = int(thisBuyNum)
            cell.buyAmount = round(cell.buyNum * self.max * thisRank)
            cell.salePrice = round(self.max * (thisRank + self.large), 2)
            cell.saleNum = int(thisBuyNum)
            cell.saleAmount = round((thisRank + self.large) * self.max * cell.saleNum)
            cell.caculateProfit()
            self.largeGrids.append(cell)
            thisRank = round(thisRank - self.large, 2)



    def generateExcel(self):
        wb = xlwt.Workbook(encoding='utf8')
        st = wb.add_sheet("网格策略")
        st.write(0, 0, '名称')
        st.write(0, 1, '档位')
        st.write(0, 2, '买入价格')
        st.write(0, 3, '买入数量')
        st.write(0, 4, '买入金额')
        st.write(0, 5, '卖出价格')
        st.write(0, 6, '卖出数量')
        st.write(0, 7, '卖出金额')
        st.write(0, 8, '手续费')
        st.write(0, 9, '利润')
        st.write(0, 10, '利润率')

        i = 0

        pp = 0
        tp = 0
        for cell in (self.smallGrids+self.middleGrids+self.largeGrids):
            cell.desc()
            style = xlwt.XFStyle()
            pattern = xlwt.Pattern()
            pattern.pattern = xlwt.Pattern.SOLID_PATTERN
            pattern.pattern_fore_colour=1
            if cell.type == '中网':
                pattern.pattern_fore_colour=51
            elif cell.type == '大网':
                pattern.pattern_fore_colour = 50
            style.pattern = pattern

            pp += cell.buyAmount
            tp += cell.profit
            i += 1
            st.write(i, 0, cell.type, style=style)
            st.write(i, 1, str(cell.rank), style=style)
            st.write(i, 2, str(cell.buyPrice), style=style)
            st.write(i, 3, str(cell.buyNum), style=style)
            st.write(i, 4, str(cell.buyAmount), style=style)
            st.write(i, 5, str(cell.salePrice), style=style)
            st.write(i, 6, str(cell.saleNum), style=style)
            st.write(i, 7, str(cell.saleAmount), style=style)
            st.write(i, 8, str(cell.serviceCharge), style=style)
            st.write(i, 9, str(cell.profit), style=style)
            st.write(i, 10, "%s%%" % (cell.profitRate * 100), style=style)

        i += 1

        st.write(i, 0, '总利润 %s' % tp)
        st.write(i, 2, '压力测试 %s' % pp)

        wb.save(self.targetDir)
        print('保存文件', self.targetDir)


