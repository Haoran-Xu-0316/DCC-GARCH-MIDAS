#未封装的函数，用于快速查看
#修改代码选择宏观不确定性，K_c和NN_c
#绘图，不保存
library(dccmidas)
require(xts)
library(readxl)
library(ggplot2)
library(openxlsx)
require(rumidas)

#Hedging Assets----------------------------------------------------------------------------------------------------------
# label <- c("CSI 300",	"中证500",	"中证1000",	"Green Bond",	"Gold",	"WTI",	"Bitcon","Bond")
label <- c("CSI 300",	"中证500",	"中证1000",	"Gold",	"WTI",	"Bitcoin","Bond")
# label=c("WTI","Gold","Silver","Palladium","Platinum","Crude Oil")
# label=c("Industrial Metals","Precious Metals",	"Energy",	"Agriculture",	"Livestock")
# label=c("WTI","Gold","Silver","Palladium","Platinum")
data <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC-data\\Hedging Assets\\hedging assets.xlsx", sheet = "All-log")
# data <- read_excel("C:\\Users\\徐浩然\\Desktop\\Commodities\\Commodity.xlsx", sheet = "log")
# data <- read_excel("C:\\Users\\徐浩然\\Desktop\\Metal\\SP\\Metals.xlsx", sheet = "Metals-SP-log")


dates <- as.Date(data$Date, format = "%Y%m%d") 
data <- xts(data[, label], order.by = dates) 
C=c(1,4,5,6,7)
# C=c(1,2,3,4,5)
# C=c(1,8,4,5,6,7)
data=data[, C]*1
label=label[C]
n=length(label)
K=12
#Macro Uncertainty--------------------------------------------------------------------------------------------------------
low <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC-data\\Macro Uncertainty\\macro uncertainty.xlsx", sheet = "All")
datess <- as.Date(low$Date, format = "%Y%m%d") 
# low <- xts(low$CPU, order.by = datess)
# low <- xts(low$EMV, order.by = datess)
# low <- xts(low$EPU, order.by = datess)
low <- xts(low$TPU, order.by = datess)
# low <- xts(low$GPR, order.by = datess)
# # low <- xts(low$GVZ, order.by = datess)
# low <- xts(low$OVX, order.by = datess)
# low <- xts(low$VIX, order.by = datess)
# low <- diff(log(low))*100
low=diff(low)
low <- low[complete.cases(low),]
#Other preparations-------------------------------------------------------------------------------------------------------



r_t <- data[, label]
# r_t <- r_t[complete.cases(r_t),]#去除缺失值
r_t=r_t[1:(nrow(r_t) - 41), ]
r_t <- head(r_t, n = nrow(r_t) - 0)
r_t <- lapply(r_t, function(x) x)

MV <- lapply(r_t, function(x) mv_into_mat(x, low, K = K, "monthly"))

K_c<-275
N_c<-22

#Modeling-------------------------------------------------------------------------------------------------------------------
dccmidas_est=dcc_fit(r_t,univ_model="DAGM_skew",distribution="norm",MV=MV,K=K,corr_model="DCCMIDAS",N_c=N_c,K_c=K_c)
# dccmidas_est=dcc_fit(r_t,univ_model="GM_skew",distribution="norm",MV=MV,K=K,corr_model="DCCMIDAS",N_c=N_c,K_c=K_c)
# dccmidas_est=dcc_fit(r_t,univ_model="GARCH",corr_model="aDCC",distribution="norm")
summary.dccmidas(dccmidas_est)


#Plot------------------------------------------------------------------------------
# plot_dccmidas(
#   dccmidas_est,
#   K_c =K_c,
#   vol_col = "black",
#   long_run_col = "red",
#   cex_axis = 0.75,
#   LWD = 1.5,
#   asset_sub = NULL
# )

#Results-------------------------------------------------------------------------------------------------------------------
# dccmidas_est


# #
# sink("C:\\Users\\徐浩然\\Desktop\\Metal\\SP\\USTPU-6.txt")
# summary.dccmidas(dccmidas_est)
# sink()




