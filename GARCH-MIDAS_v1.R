#为了获取单个Double Asymmetry Garch Midas 的短期波动和长期波动
#不一定会使用
require(xts)
library(readxl)
library(ggplot2)
library(openxlsx)
require(rumidas)


#Hedging Assets----------------------------------------------------------------------------------------------------------
label <- c("CSI 300",	"中证500",	"中证1000",	"Green Bond",	"Gold",	"WTI",	"Bitcon","Bond")
data <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC data\\Hedging Assets\\hedging assets.xlsx", sheet = "All-log100")
df_l=data[-1,1]
df_s=data[-1,1]

dates <- as.Date(data$Date, format = "%Y%m%d") 
data <- xts(data[, label], order.by = dates) 
C=c(1,4,5,6,7)
data=data[, C]
label=label[C]
n=length(label)
#Macro Uncertainty--------------------------------------------------------------------------------------------------------
low <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC data\\Macro Uncertainty\\macro uncertainty.xlsx", sheet = "All")
datess <- as.Date(low$Date, format = "%Y%m%d") 

# low <- xts(low$GPR, order.by = datess)
# low <- xts(low$EPU, order.by = datess)
# low <- xts(low$CEPU, order.by = datess)
# low <- xts(low$CPU, order.by = datess)
# low <- xts(low$EMV, order.by = datess)
# low <- xts(low$TPU, order.by = datess)
# low <- low[-c(1:50), ]
#Other preparations-------------------------------------------------------------------------------------------------------
r_t <- data[, label]
r_t <- r_t[complete.cases(r_t),]#去除缺失值



Vol=function(asset,macro,K){
  
  
r_t = r_t[,asset]
low <- xts(low[,macro], order.by = datess)
low=diff(low)

MV<-mv_into_mat(r_t,low,K,"monthly")
fit<-ugmfit(model="GM",skew="YES",distribution="norm",r_t,MV,K)

summary.rumidas(fit)
fit["loglik"]
fit["inf_criteria"]
fit["est_in_s"]

est=unlist(fit["rob_coef_mat"])

alpha=est["rob_coef_mat.Estimate1"]
gamma=est["rob_coef_mat.Estimate2"]
beta=est["rob_coef_mat.Estimate3"]
m=est["rob_coef_mat.Estimate4"]
theta=est["rob_coef_mat.Estimate5"]
w2=est["rob_coef_mat.Estimate6"]

param<-c(alpha,beta,gamma,m,theta,w2)

SV=GM_cond_vol(param,r_t,MV,K)
LV=GM_long_run_vol(param, r_t, MV, K)

SVLV <- list(SV,LV,fit["est_in_s"])
return(SVLV)
}
V=Vol("Green Bond","EMV",24)
S = V[[1]]
L =V[[2]]

plot(S)
lines(L, col = "red",lwd = 2)









