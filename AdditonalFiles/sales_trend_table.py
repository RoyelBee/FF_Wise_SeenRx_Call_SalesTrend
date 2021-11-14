import xlrd
import AdditonalFiles.functions as fn

def all_Sales_achiv_data():
    xl = xlrd.open_workbook(r"./Data/SalesAchivment/sales_achiv.xlsx")
    sh = xl.sheet_by_name('Sheet1')
    th = ""
    td = ""
    # Select all columns header name and placed All name in serial
    for i in range(0, 1):
        th = th + "<tr>\n"
        # th = th + "<th class=\"unit\">ID</th>"

        for j in range(0, sh.ncols):
            th = th + "<th class=\"unit\">" + str(sh.cell_value(i, j)) + "</th>\n"
        th = th + "</tr>\n"

    # Now placed all data
    for i in range(1, sh.nrows):
        # td = td + "<tr>\n"
        # td = td + "<td class=\"unit\">" + str(i) + "</td>"

        for j in range(0, 1):
            td = td + "<td id=\"unit\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + "</td>\n"

        for j in range(1, 2):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(2, 3, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(3, 4, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(4, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        td = td + "</tr>\n"
    html = th + td
    print('3. Sales Achievement Table Created')
    return html

def all_Sales_trend_table_data():
    xl = xlrd.open_workbook(r"./Data/SalesTrend/sales_trend_data.xlsx")
    sh = xl.sheet_by_name('Sheet1')
    th = ""
    td = ""
    # Select all columns header name and placed All name in serial
    for i in range(0, 1):
        th = th + "<tr>\n"
        # th = th + "<th class=\"unit\">ID</th>"

        for j in range(0, sh.ncols):
            th = th + "<th class=\"unit\">" + str(sh.cell_value(i, j)) + "</th>\n"
        th = th + "</tr>\n"

    # Now placed all data
    for i in range(1, sh.nrows):
        # td = td + "<tr>\n"
        # td = td + "<td class=\"unit\">" + str(i) + "</td>"

        for j in range(0, 1):
            td = td + "<td id=\"unit\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + "</td>\n"

        for j in range(1, 2):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(2, 3, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(3, 4, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(4, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        td = td + "</tr>\n"
    html = th + td
    print('4. Sales Trend Table Created')

    return html

def rsm_Sales_table_data(rsm):
    xl = xlrd.open_workbook(r"./Data/SalesTrend/SalesTrend_" + str(rsm) + ".xlsx")
    sh = xl.sheet_by_name('Sheet1')
    th = ""
    td = ""
    # Select all columns header name and placed All name in serial
    for i in range(0, 1):
        th = th + "<tr>\n"
        # th = th + "<th class=\"unit\">ID</th>"

        for j in range(0, sh.ncols):
            th = th + "<th class=\"unit\">" + str(sh.cell_value(i, j)) + "</th>\n"
        th = th + "</tr>\n"

    # Now placed all data
    for i in range(1, sh.nrows):
        # td = td + "<tr>\n"
        # td = td + "<td class=\"unit\">" + str(i) + "</td>"

        for j in range(0, 1):
            td = td + "<td id=\"unit\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + "</td>\n"

        for j in range(1, 2):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(2, 3, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(3, 4, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(4, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        td = td + "</tr>\n"
    html = th + td
    print('5. SeenRx Table Created')

    return html

def rsm_Sales_achiv_data(rsm):
    xl = xlrd.open_workbook(r"./Data/SalesAchivment/SalesAchivment_"+str(rsm)+".xlsx")
    sh = xl.sheet_by_name('Sheet1')
    th = ""
    td = ""
    # Select all columns header name and placed All name in serial
    for i in range(0, 1):
        th = th + "<tr>\n"
        # th = th + "<th class=\"unit\">ID</th>"

        for j in range(0, sh.ncols):
            th = th + "<th class=\"unit\">" + str(sh.cell_value(i, j)) + "</th>\n"
        th = th + "</tr>\n"

    # Now placed all data
    for i in range(1, sh.nrows):
        # td = td + "<tr>\n"
        # td = td + "<td class=\"unit\">" + str(i) + "</td>"

        for j in range(0, 1):
            td = td + "<td id=\"unit\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + "</td>\n"

        for j in range(1, 2):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(2, 3, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(3, 4, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        for j in range(4, sh.ncols):
            td = td + "<td id=\"idcol\" style=\"background-color:" + \
                 str(fn.warning(str((sh.cell_value(i, 0))))) + "\">" + str((sh.cell_value(i, j))) + '%' + "</td>\n"
        td = td + "</tr>\n"

    html = th + td
    print('6. Doctor Call Table Created')

    return html
