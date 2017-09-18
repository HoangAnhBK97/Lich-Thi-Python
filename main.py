import urllib
import xlrd
import pdftables_api
import os

url = raw_input("Nhap URL cua lich thi : ")
try :
    print "Dang tai ve danh sach thi ..."
    urllib.urlretrieve(url,"lichthi.pdf")
    print "Da tai ve thanh cong ! Dang chuyen doi du lieu ..."
    try :
        #Lay Key API : https://pdftables.com/pdf-to-excel-api
        c = pdftables_api.Client('gfqvfuza72jl')
        c.xlsx('lichthi.pdf', 'lichthi.xlsx')

        print " "*30,"Tra cuu lich thi"

        file_location =  "lichthi.xlsx"
        wb = xlrd.open_workbook(file_location)
        sheet = wb.sheet_by_index(0)

        list_hoc_phan = raw_input("Nhap cac ma hoc phan muon tra cuu lich thi (Ngan cach nhau boi 'space') : ")
        try :
            hp = list_hoc_phan.split()
            data = [int(x) for x in hp]
            print " "*30,"Lich thi cua ban"
            count = 0
            for rows in range(sheet.nrows):
                if sheet.cell_value(rows,0) in data :
                    count += 1
                    for cols in range(sheet.ncols):
                        print sheet.cell_value(rows,cols),
                    print "\n"
            if count == 0 :
                print "Khong co hoc phan nao trong lich thi"
        except Exception :
            print "Hoc phan ban nhap khong dung !! Vui long thu lai "
        os.remove('lichthi.pdf')
        os.remove('lichthi.xlsx')
    except Exception :
        print "Khong chuyen doi duoc File do key API da het han !!"
except Exception  :
    print "Loi URL ! Ban phai nhap URL cua lich thi !! Len trang http://dtdh.hust.edu.vn de xem URL nhe ^^"
