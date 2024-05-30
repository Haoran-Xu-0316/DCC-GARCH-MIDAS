
library(rumidas)
require(xts)
library(readxl)
library(ggplot2)
library(openxlsx)
require(rumidas)

#Hedging Assets----------------------------------------------------------------------------------------------------------
label <- c("CSI 300","S&P 500","Gold","WTI","Bitcoin","Bond")

data <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC-data-China\\Hedging Assets\\hedging assets.xlsx", sheet = "All-log")
dates <- as.Date(data$Date, format = "%Y%m%d") 
data <- xts(data[, label], order.by = dates) 
C=c(1,2,3,4,5,6)

data=data[, C]*1
label=label[C]


#Macro Uncertainty--------------------------------------------------------------------------------------------------------
low <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC-data-China\\Macro Uncertainty\\macro uncertainty.xlsx", sheet = "All")
datess <- as.Date(low$Date, format = "%Y%m%d") 
lowname=colnames(low)[2:6]

low=low[-1]
low <- xts(low, order.by = datess)

low=diff(log(low))
low_d <- low[complete.cases(low),]

low=low_d[,5]
# "CPU" "EMV" "EPU" "TPU" "GPR"
#Other preparations-------------------------------------------------------------------------------------------------------
r_t <- data[, label]
r_t <- r_t[complete.cases(r_t),]#去除缺失值
r_t=r_t[1:(nrow(r_t) - 30), ]

# r_t=r_t$`CSI 300`
# r_t=r_t$`S&P 500`
r_t=r_t$`Gold`
# r_t=r_t$WTI
# r_t=r_t$Bitcoin
# r_t=r_t$Bond

LL <- list()
AIC <- list()
BIC <- list()

for (K in 13:30) {
print(paste("+++++++++++ K =", K,"+++++++++++"))
mv_m<-mv_into_mat(r_t,low,K=K,"monthly")
fit<-ugmfit(model="DAGM",skew="YES",distribution="norm",r_t,mv_m,K=K)
summary.rumidas(fit)
print(fit["loglik"])
print(fit["inf_criteria"])
LL <- append(LL, fit["loglik"][[1]])
# AIC <- append(AIC, fit["inf_criteria"][[1]])
# BIC <- append(BIC, fit["inf_criteria"][[2]])
print("---------------------------------------------------")
}

# 
# print("#################################################################################################")
# print(paste("LL: The best model is when K =",which.max(LL)+1,"and equals =",LL[[which.max(LL)]]))
# print("#################################################################################################")
# K=which.max(LL)+1
# mv_m<-mv_into_mat(r_t,low,K=K,"monthly")
# fit<-ugmfit(model="DAGM",skew="YES",distribution="norm",r_t,mv_m,K=K)
# summary.rumidas(fit)
# print(fit["loglik"])
# print(fit["inf_criteria"])



# est_val<-c(alpha=0.1750,gamma=-0.1731,beta=0.9106 ,m=-5.2155,theta_pos=0.0431,w2_pos=1.0381,theta_neg=0.0024,w2_neg=2.3311)
# vol<-DAGM_cond_vol(est_val,r_t,mv_m,K=11)
# voll=DAGM_long_run_vol(est_val,r_t,mv_m,K=11,lag_fun = "Beta")
# 
# plot(vol)
# lines(voll, col = "red",lwd = 2)


# 
# sink("C:\\Users\\徐浩然\\Desktop\\Results\\S&P500-CPU.txt")
# for (K in 25) {
#   print(paste("+++++++++++ K =", K,"+++++++++++"))
#   mv_m<-mv_into_mat(r_t,low,K=K,"monthly")
#   fit<-ugmfit(model="DAGM",skew="YES",distribution="norm",r_t,mv_m,K=K)
#   summary.rumidas(fit)
#   print(fit["loglik"])
#   print(fit["inf_criteria"])
#   LL <- append(LL, fit["loglik"][[1]])
#   # AIC <- append(AIC, fit["inf_criteria"][[1]])
#   # BIC <- append(BIC, fit["inf_criteria"][[2]])
#   print("---------------------------------------------------")
# }
# sink()


