# Рассылка отчетов #
start outlook.exe
$Outlook = New-Object -ComObject Outlook.Application
$username = $env:UserName
$path = "C:\Users\$username\Documents\Reports"
$files = Get-ChildItem -Path C:\Users\$username\Documents\Reports -Name # Тут можно унифицировать выборку по формату файла и пр.
$time = [DateTime]::Today.AddDays(-1).ToString("dd-MM-yyyy") # В этом случае -1 день, потому что отчет делается за прошедший день
$Mail = $Outlook.CreateItem(0)
$Mail.To = "" # Сюда в кавычки нужно вставить получателей в формате blabla@mail.ru; blablabla@mail.ru;
$Mail.CC = "" # Сюда вставляем в аналогичном формате адреса пользователей, кому направляется копия письма
$Mail.Subject = "Отчет за $time"
$Mail.Body = "Файлы во вложении"
foreach ($a in $files) {
    $file = $a
    $attachment = "$path\$file"
    $Mail.Attachments.Add($attachment)
}
$Mail.Send()
