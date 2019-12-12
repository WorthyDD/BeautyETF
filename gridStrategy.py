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

    def desc(self):
        print self.type, self.rank, self.buyPrice, self.buyNum, self.buyAmount, self.salePrice, self.saleNum, self.saleAmount, self.profit, self.profitRate




# 目标最高价
max = 0.80

# 大网
large = 0.3

# 中网
middle = 0.15

# 小网
small = 0.05

# 压力测试价
pressPrice = 0

# 每一小格递增买入比率
stepRate = 0.15

# 从哪一个档位开始递增投资
initalStep = 2

min = 0.4 * max

# 初始买入数量
minBuyNum = 1000

smallGrids = []
middleGrids = []
largeGrids = []


def makeGrids():
    # 档位对应买入数量
    rateNum = {}

    # 小网织网

    thisRank = 1.0
    index = 1
    thisBuyNum = minBuyNum
    while thisRank >= 0.4:
        cell = GridCell()
        cell.rank = thisRank
        cell.type = "小网"
        if index > initalStep:
            thisBuyNum = round((1 + stepRate) * thisBuyNum)
        cell.buyPrice = round(max * thisRank, 2)
        cell.buyNum = thisBuyNum
        cell.buyAmount = round(cell.buyNum * max * thisRank)
        cell.salePrice = round(max * (thisRank + small), 2)
        cell.saleNum = thisBuyNum
        cell.saleAmount = round((thisRank + small) * max * cell.saleNum)
        cell.profit = round(cell.saleAmount - cell.buyAmount)
        cell.profitRate = round(cell.profit / cell.buyAmount, 2)
        smallGrids.append(cell)
        index += 1
        thisRank = round(thisRank - small, 2)
        rateNum[thisRank] = thisBuyNum
        print thisRank

    print "rateNum", rateNum

    # 中网织网
    thisRank = round(1.0 - middle, 2)
    while thisRank >= 0.4:
        thisBuyNum = rateNum[thisRank]
        cell = GridCell()
        cell.rank = thisRank
        cell.type = "中网"
        cell.buyPrice = round(max * thisRank, 2)
        cell.buyNum = thisBuyNum
        cell.buyAmount = round(cell.buyNum * max * thisRank)
        cell.salePrice = round(max * (thisRank + middle), 2)
        cell.saleNum = thisBuyNum
        cell.saleAmount = round((thisRank + middle) * max * cell.saleNum)
        cell.profit = round(cell.saleAmount - cell.buyAmount)
        cell.profitRate = round(cell.profit / cell.buyAmount, 2)
        middleGrids.append(cell)
        thisRank = round(thisRank - middle, 2)

    # 大网织网
    thisRank = round(1.0 - large, 2)
    while thisRank >= 0.4:
        thisBuyNum = rateNum[thisRank]
        cell = GridCell()
        cell.rank = thisRank
        cell.type = "大网"
        cell.buyPrice = round(max * thisRank, 2)
        cell.buyNum = thisBuyNum
        cell.buyAmount = round(cell.buyNum * max * thisRank)
        cell.salePrice = round(max * (thisRank + large), 2)
        cell.saleNum = thisBuyNum
        cell.saleAmount = round((thisRank + large) * max * cell.saleNum)
        cell.profit = round(cell.saleAmount - cell.buyAmount)
        cell.profitRate = round(cell.profit / cell.buyAmount, 2)
        largeGrids.append(cell)
        thisRank = round(thisRank - large, 2)



def generateExcel():
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
    st.write(0, 8, '利润')
    st.write(0, 9, '利润率')

    i = 0

    pp = 0
    tp = 0
    for cell in (smallGrids+middleGrids+largeGrids):
        cell.desc()
        pp += cell.buyAmount
        tp += cell.profit
        i += 1
        st.write(i, 0, cell.type)
        st.write(i, 1, str(cell.rank))
        st.write(i, 2, str(cell.buyPrice))
        st.write(i, 3, str(cell.buyNum))
        st.write(i, 4, str(cell.buyAmount))
        st.write(i, 5, str(cell.salePrice))
        st.write(i, 6, str(cell.saleNum))
        st.write(i, 7, str(cell.saleAmount))
        st.write(i, 8, str(cell.profit))
        st.write(i, 9, "%s%%" % (cell.profitRate * 100))
    i += 1

    st.write(i, 0, '总利润 %s' % tp)
    st.write(i, 2, '压力测试 %s' % pp)

    wb.save('/Users/wuxi/Desktop/grid.xlsx')


if __name__ == '__main__':
    makeGrids()
    print '完成织网'
    generateExcel()