#-*- coding: sjis -*-
#!/usr/bin/ruby -Ks

require 'win32ole'

def show(msg, title)
  wsh = WIN32OLE.new('WScript.Shell')
  wsh.Popup(msg, 0, title, 0 + 64 + 0x40000)
end

app = WIN32OLE.new('Excel.Application')
book = app.Workbooks.Open(app.GetOpenFilename('Microsoft Excelブック,*.xls'))

begin
  book.Activate
  book.Worksheets.count.downto(1){|sheetnum|
    sheet = book.Worksheets(sheetnum)
    sheet.Activate
    sheet.Range("A1").Activate
  }

  book.save
ensure
  book.close(false)
  app.quit
end

show("できたよ！！", "message")

