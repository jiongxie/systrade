import win32com.client

instCpMarketWatch = win32com.client.Dispatch("CpSysDib.CpMarketWatch")


instCpMarketWatch.SetInputValue(',1)

value = instCpMarketWatch.GetHeaderValue(1)

