
options(digits = 3)
library(rumidas)
require(xts)
library(readxl)
library(ggplot2)
library(openxlsx)
require(rumidas)

library(FinTS)
library(tseries)
library(forecast)
library(stats)


A <- read_excel("C:\\Users\\徐浩然\\Desktop\\DS.xlsx", sheet = "All-log")

M <- read_excel("C:\\Users\\徐浩然\\Desktop\\DS.xlsx", sheet = "M")


adf.test(A$`CSI 300`)
adf.test(A$`S&P 500`)
adf.test(A$Bond)
adf.test(A$Gold)
adf.test(A$WTI)
adf.test(A$Bitcoin)

adf.test(M$CPU)
adf.test(M$EMV)
adf.test(M$EPU)
adf.test(M$TPU)
adf.test(M$GPR)

l=15
b1=Box.test(A$`CSI 300`, lag = l, type = "Ljung-Box")
Box.test(A$`S&P 500`, lag = l, type = "Ljung-Box")
Box.test(A$Bond, lag = l, type = "Ljung-Box")
Box.test(A$Gold, lag = l, type = "Ljung-Box")
Box.test(A$WTI, lag = l, type = "Ljung-Box")
Box.test(A$Bitcoin, lag = l, type = "Ljung-Box")

Box.test(M$CPU, lag = l, type = "Ljung-Box")
Box.test(M$EMV, lag = l, type = "Ljung-Box")
Box.test(M$EPU, lag = l, type = "Ljung-Box")
Box.test(M$TPU, lag = l, type = "Ljung-Box")
Box.test(M$GPR, lag = l, type = "Ljung-Box")

lb_test_result <- LjungBoxTest(ts_data, lag = 10)

ll=10
a1=ArchTest(A$`CSI 300`, lags=ll) 
a2ArchTest(A$`S&P 500`, lags=ll) 
ArchTest(A$Bond, lags=ll) 
ArchTest(A$Gold, lags=ll) 
ArchTest(A$WTI, lags=ll) 
ArchTest(A$Bitcoin, lags=ll)

library(Hmisc)#加载包
sink("C:\\Users\\徐浩然\\Desktop\\cor.txt")
res <- rcorr(as.matrix(A[2:7]))
res[["r"]]
res
sink()



x <- rnorm (100)
Box.test (x, lag = 1)
Box.test (x, lag = 1, type = "Ljung")
