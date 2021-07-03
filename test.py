# 1. data_file 第1列建议是索引列
# 2. filling_loc 大写金额名，必须与 addcnamount 参数 cn_colname 相同
# 3. filling_loc 的 value 必须与 data_file 中的列名相同
# 4. pagelayout 的参数 margins, 给定的值是 inch 单位， not cm。
from yolib import pattern_copier

pattern_file = 'temp.xlsx'
data_file = '请款单明细表.xlsx'
out_file = 'temp1.xlsx'
filling_loc = {
    'M2':'填单日期',
    'E3':'收款单位', 'M3':'请款部门',
    'E4':'银行账号', 'M4':'开户银行',
    'F5':'金额（大写）',
    'M5':'转款金额', 
    'B7':'用途', 
    'O9':'单据数量'
}

copier = pattern_copier(
    pattern_file,
    data_file,
    out_file,
    filling_loc
)

copier.addcnamount('金额（大写）','转款金额')
copier.filterdatas('填单日期',['2021-6-23','2021-6-24'])
copier.copyblocks(number_of_blocks_per_page=1)
margins={
        'top':0.4, 'bottom':0.35,
        'left':0.25, 'right':0.25,
        'header':0.3, 'footer':0.3
    }
copier.pagelayout(margins=margins, orientation='landscape', scale=67)
copier.output()

