#Using namespace System.Threading
#Using namespace System.Threading.Tasks


$Excel = New-Object -ComObject Excel.Application
$Excel.Caption = "Host of ZKL"
$Excel.Visible = $true                           # Видимость приложения включена
$WorkBook = $Excel.Workbooks.Add()               # Добавляю книгу в ексель
$HostInfo = $WorkBook.Worksheets.Item(1)         # Привязка к первому листу
$HostInfo.Name = "Attribute of AD Comps"                # Имя листа
[int]$Row = 1 # Ряды
[int]$Col = 1 # Столбцы

$HostInfo.Range('A1').ColumnWidth = 25  # Ширина колонок
$HostInfo.Range('B1').ColumnWidth = 25
$HostInfo.Range('C1').ColumnWidth = 25
$HostInfo.Range('D1').ColumnWidth = 25
$HostInfo.Range('E1').ColumnWidth = 25
$HostInfo.Range('F1').ColumnWidth = 25

$HostInfo.Cells.Item($Row,$Col).Font.Size = 12    # Размер шрифта
$HostInfo.Cells.Item($Row,$Col).Font.Bold = $true # Тип шрифта
$HostInfo.Cells.Item($Row,$Col) = "Host name"     # Заголовки колонок

$HostInfo.Cells.Item($Row,$Col+1).Font.Size = 12
$HostInfo.Cells.Item($Row,$Col+1).Font.Bold = $true
$HostInfo.Cells.Item($Row,$Col+1) = "User"

$HostInfo.Cells.Item($Row,$Col+2).Font.Size = 12
$HostInfo.Cells.Item($Row,$Col+2).Font.Bold = $true
$HostInfo.Cells.Item($Row,$Col+2) = "User CN"

$HostInfo.Cells.Item($Row,$Col+3).Font.Size = 12
$HostInfo.Cells.Item($Row,$Col+3).Font.Bold = $true
$HostInfo.Cells.Item($Row,$Col+3) = "Primary User"

$HostInfo.Cells.Item($Row,$Col+4).Font.Size = 12
$HostInfo.Cells.Item($Row,$Col+4).Font.Bold = $true
$HostInfo.Cells.Item($Row,$Col+4) = "extensionAttribute1"

$HostInfo.Cells.Item($Row,$Col+5).Font.Size = 12
$HostInfo.Cells.Item($Row,$Col+5).Font.Bold = $true
$HostInfo.Cells.Item($Row,$Col+5) = "extensionAttribute2"


$xml = Import-Clixml -Path "C:\Users\stepa\Desktop\ExtAttr\xml1.xml" # Получаю хэш-таблицу со значинем (имя => {пароль, дата})

        $itr = 1
          foreach($HostName in $xml.Keys)                # Перебираю все ключи в массиве 
            { 
              $HostInfo.Cells.Item($Row+$itr,$Col) = $HostName  # Пишу в ячейку имя хоста
              $HostInfo.Cells.Item($Row+$itr,$Col+1) = $xml[$HostName].Item(0) # Пишу в ячейку имя юзера
              $HostInfo.Cells.Item($Row+$itr,$Col+2) = $xml[$HostName].Item(1) # Пишу в ячейку дату и время
              $HostInfo.Cells.Item($Row+$itr,$Col+3) = $xml[$HostName].Item(2) # Пишу в PrimaryUserName
              $HostInfo.Cells.Item($Row+$itr,$Col+4) = $xml[$HostName].Item(3) # Пишу в ячейку extensionAttribute1
              $HostInfo.Cells.Item($Row+$itr,$Col+5) = $xml[$HostName].Item(4) # Пишу в ячейку extensionAttribute2
              $itr++
            }
<#
        $itr = 1
          foreach($UserName in $xml.Values.Item(0))        # Получаю все значения паролей в массиве
            { 
              $HostInfo.Cells.Item($Row+$itr,$Col+1) = $UserName   # Пишу в ячейку имя юзера                                          
              $itr++
            } 

        $itr = 1
          foreach($UserCN in $xml.Values.Item(1))        # Получаю все значения даты и времени истекания пароля
            { 
              $HostInfo.Cells.Item($Row+$itr,$Col+2) = $UserCN # Пишу в ячейку дату и время
              $itr++
            }

        $itr = 1
          foreach($PrimaryUser in $xml.Values.Item(2))        # Получаю все значения PrimaryUser
            { 
              $HostInfo.Cells.Item($Row+$itr,$Col+3) = $PrimaryUser # Пишу в PrimaryUserName
              $itr++
            }

        $itr = 1
          foreach($extensionAttribute1 in $xml.Values.Item(3))        # Получаю все значения даты и времени истекания пароля
            { 
              $HostInfo.Cells.Item($Row+$itr,$Col+4) = $extensionAttribute1  # Пишу в ячейку extensionAttribute1
              $itr++
            }

        $itr = 1
          foreach($extensionAttribute2 in $xml.Values.Item(4))        # Получаю все значения даты и времени истекания пароля
            { 
              $HostInfo.Cells.Item($Row+$itr,$Col+5) = $extensionAttribute2  # Пишу в ячейку extensionAttribute2
              $itr++
            }
#>
