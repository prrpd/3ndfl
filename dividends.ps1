<#
подготовка
заказать xlsx файл с дивами в ТП Тинкова
удалить лишние данные, изменить шапку
проверить что нет лишних таблиц в csv
для американских деп расписок типа МТС изменить страну с России на США


$workingPath = pwd
if (($env:Path -split ';') -notcontains $workingPath) {
    $env:Path += ";$workingPath"
}
Add-Type -Path "$($workingPath)\WebDriver.dll"
#или можно без изменения переменной path сразу в add-type указать путь к WebDriver.dll
#>

<#
особенности - галка определять курс автоматически стояла, но курс на дату уплаты налога не подтянулся. На сайте мне это было показано, поэтому галку отжал и снова нажал - курс появился
#>

<#
аналоги
https://github.com/gerasiov/3ndfl/blob/main/nalog.ipynb
#>


$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver
$ChromeDriver.Navigate().GoToURL('https://lkfl2.nalog.ru/lkfl/login')
$BaseIDpath = 'Ndfl3Package.payload.sheetB.sources'
$records = import-csv "C:\dividends.csv" -Delimiter ';'

foreach ($record in $records) {
	$Index = $records.IndexOf($record)
	
	write-output 'Добавление источника дохода'
	$ChromeDriver.FindElement([OpenQA.Selenium.By]::ClassName('IncomeSources_addButton__1jhpg')).click()
	start-sleep -Milliseconds 100
	
	write-output "Источник дохода $($Index+1) - $($record.Name)"
	try {
		#номер react-tabs-3 может меняться, поэтому если не сработало нажатие на источник дохода, то нужно проверить/изменить путь к элементу
		$ChromeDriver.FindElement([OpenQA.Selenium.By]::CssSelector("#react-tabs-3 > div > div:nth-child($($Index+1)) > div.Spoiler_spoilerItemHeader__1RM7f > div.Spoiler_title__PVXtF > div > div.IncomeSources_title__2RplA > div > span")).click()
	} catch {
		throw 'Не найден элемент Источник дохода'
	}
	write-output 'Наименование'	
	$incomeSourceName=$BaseIDpath+"[$index].incomeSourceName"
	$ChromeDriver.FindElement([OpenQA.Selenium.By]::ID($incomeSourceName)).Clear()
	$ChromeDriver.FindElement([OpenQA.Selenium.By]::ID($incomeSourceName)).SendKeys("$($record.Name)")
	
	write-output 'Страна источника выплаты'
	$oksmIst=$BaseIDpath+"[$index].oksmIst"
	switch ($record.Country) {
		'США' 				{$country = 840; break}
		'Великобритания'	{$country = 826; break}
		'Германия' 			{$country = 276; break}
		'Кипр' 				{$country = 196; break}
		'Тайвань' 			{$country = 158; break}
		'ДЖЕРСИ' 			{$country = 832; break}
		#'РОССИЯ' 			{$country = 196; break}
	}
	$element = $ChromeDriver.FindElement([OpenQA.Selenium.By]::ID($oksmIst))
	$inputTag = $element.FindElement([OpenQA.Selenium.By]::TagName('input'))
	$inputTag.SendKeys("$country")
	$inputTag.SendKeys([OpenQA.Selenium.Keys]::Enter)
	
	write-output 'Страна зачисления выплаты'
	$oksmZach=$BaseIDpath+"[$index].oksmZach"
	$element = $ChromeDriver.FindElement([OpenQA.Selenium.By]::ID($oksmZach))
	$inputTag = $element.FindElement([OpenQA.Selenium.By]::TagName('input'))
	$inputTag.SendKeys('643')
	$inputTag.SendKeys([OpenQA.Selenium.Keys]::Enter)
	
	write-output 'Код дохода'
	$incomeTypeCode = $BaseIDpath+"[$index].incomeTypeCode"
	$element = $ChromeDriver.FindElement([OpenQA.Selenium.By]::ID($incomeTypeCode))
	$inputTag = $element.FindElement([OpenQA.Selenium.By]::TagName('input'))
	$inputTag.SendKeys('1010')
	$inputTag.SendKeys([OpenQA.Selenium.Keys]::Enter)
	
	write-output 'Предоставить нал-й вычет'
	$taxDeductionCode = $BaseIDpath+"[$index].taxDeductionCode"
	$element = $ChromeDriver.FindElement([OpenQA.Selenium.By]::ID($taxDeductionCode))
	$inputTag = $element.FindElement([OpenQA.Selenium.By]::TagName('input'))
	$inputTag.SendKeys('не')
	$inputTag.SendKeys([OpenQA.Selenium.Keys]::Enter)
	
	write-output 'Сумма дохода в валюте'	
	$incomeAmountCurrency = $BaseIDpath+"[$index].incomeAmountCurrency"
	$SumBeforeTaxReplaceComma = $record.SumBeforeTax -replace (',','.')
	$ChromeDriver.FindElement([OpenQA.Selenium.By]::ID($incomeAmountCurrency)).SendKeys("$SumBeforeTaxReplaceComma")
	
	write-output 'Дата получения дохода'
	$incomeDate = $BaseIDpath+"[$index].incomeDate"
	$element = $ChromeDriver.FindElement([OpenQA.Selenium.By]::ID($incomeDate))
	$inputTag = $element.FindElement([OpenQA.Selenium.By]::TagName('input'))
	$inputTag.SendKeys("$($record.PaymentDay)")
	$inputTag.SendKeys([OpenQA.Selenium.Keys]::Enter)
	
	if ($record.TaxedByForeignAgent -ne '0,00') {
		write-output 'Дата уплаты налога'
		$taxPaymentDate = $BaseIDpath+"[$index].taxPaymentDate"
		$element = $ChromeDriver.FindElement([OpenQA.Selenium.By]::ID($taxPaymentDate))
		$inputTag = $element.FindElement([OpenQA.Selenium.By]::TagName('input'))
		$inputTag.SendKeys("$($record.PaymentDay)")
		$inputTag.SendKeys([OpenQA.Selenium.Keys]::Enter)
		
		write-output 'Сумма налога в иностранной валюте'
		$paymentAmountCurrency = $BaseIDpath+"[$index].paymentAmountCurrency"
		$TaxedByForeignAgent = $record.TaxedByForeignAgent -replace (',','.')
		$ChromeDriver.FindElement([OpenQA.Selenium.By]::ID($paymentAmountCurrency)).SendKeys("$TaxedByForeignAgent")
	}
	
	write-output 'Наименование валюты'
	switch ($record.currency) {
		'USD' {$currency = 840; break}
		'EUR' {$currency = 978; break}
	}
	$currencyCode = $BaseIDpath+"[$index].currencyCode"
	$element = $ChromeDriver.FindElement([OpenQA.Selenium.By]::ID($currencyCode))
	$inputTag = $element.FindElement([OpenQA.Selenium.By]::TagName('input'))
	$inputTag.SendKeys("$currency")
	$inputTag.SendKeys([OpenQA.Selenium.Keys]::Enter)
	
	write-output 'Определить курс автоматически'
	#$ChromeDriver.FindElement([OpenQA.Selenium.By]::ClassName('jq-checkbox')).click()
	#wait for checkbox to appear
	Start-sleep -Milliseconds 350
	$ChromeDriver.FindElement([OpenQA.Selenium.By]::ID("checkbox_$($index+2)")).SendKeys([OpenQA.Selenium.Keys]::Space)
	
	Write-Output '-----'
	Write-Output "Завершён ввод данных строки $($Index+1)"
	Write-Output "Компании $($record.Name)"
	Write-Output ' '
}