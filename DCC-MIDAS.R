#未封装的函数，用于快速查看
#修改代码选择宏观不确定性，K，K_c和N_c
#获取相关系数

library(dccmidas)
require(xts)
library(readxl)
library(ggplot2)
library(openxlsx)
require(rumidas)
library(Rcpp)
# 
trace(dccmidas:::dcc_fit, edit = T) # 修改后保存
# trace(dccmidas:::summary.dccmidas, edit = T) 

# source('C:\\Users\\徐浩然\\Desktop\\UR\\R file_a\\functions.R')
# 
# 
# source('C:\\Users\\徐浩然\\Desktop\\github\\dccmidas-master\\R\\RcppExports.R')
# source('C:\\Users\\徐浩然\\Desktop\\github\\dccmidas-master\\R\\data.R')
# Rcpp::sourceCpp('C:\\Users\\徐浩然\\Desktop\\github\\dccmidas-master\\src\\RcppExports.cpp')
#Hedging Assets----------------------------------------------------------------------------------------------------------

label <- c("CSI 300","S&P 500","Gold","WTI","Bitcoin","Bond")

data <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC-data-China\\Hedging Assets\\hedging assets.xlsx", sheet = "All-log")
d=data
dates <- as.Date(data$Date, format = "%Y%m%d") 
data <- xts(data[, label], order.by = dates) 
# C=c(1,3,4,5,6)# China
# C=c(2,3,4,5,6) # US
C=c(1,2,3,4,5,6)# U

data=data[, C]*1
label=label[C]
n=length(label)

K=12
K_c<-220
N_c<-5

#Macro Uncertainty--------------------------------------------------------------------------------------------------------
low <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC-data-China\\Macro Uncertainty\\macro uncertainty.xlsx", sheet = "All")
datess <- as.Date(low$Date, format = "%Y%m%d") 
lowname=colnames(low)[2:6]

low=low[-1]
low <- xts(low, order.by = datess)

low=diff(log(low))
low_d <- low[complete.cases(low),]
# sink("C:\\Users\\徐浩然\\Desktop\\DCC-data-China\\DCCMIDAS220.txt")
# sink("C:\\Users\\徐浩然\\Desktop\\DCC-data-US\\DCCMIDAS220.txt")
# sink("C:\\Users\\徐浩然\\Desktop\\DCC-data-U\\DCCMIDAS180.txt")
# for (z in 1:length(lowname)){
  low=low_d[,5]
  ###### functions-------------------------------------------------------------------
  
#   ## 相关系数
#   Cor=function(df,what){
#     for (i in 1:n) {
#       if (i+1 <= n) {
#         for (j in (i+1):n){
#           # print(i)
#           # print(j)
#           r=as.numeric(dccmidas_est[[what]][i,j,])
#           r <- cbind(data[c(1:length(r)),1],r)
#           r <- r[,-1]
#           # print(length(r))
#           df=df[c(1:length(r)),]
#           df <- cbind(df, r[,1])
#         }
#       }
#     }
#     
#     df <- df[-c(1:K_c), ]
#     return(df)
#   }
#   ## 协方差
#   Cov=function(df,what){
#     for (i in 1:n) {
#       if (i+1 <= n) {
#         for (j in (i+1):n){
#           # print(i)
#           # print(j)
#           r=as.numeric(dccmidas_est[[what]][i,j,])
#           r <- cbind(data[c(1:length(r)),1],r)
#           r <- r[,-1]
#           # print(length(r))
#           df=df[c(1:length(r)),]
#           df <- cbind(df, r[,1])
#         }
#       }
#     }
#     df <- df[-c(1:K_c), ]
#     return(df)
#   }
#   
#   ## 方差
#   Var=function(df,what){
#     for (i in 1:n) {
#       # print(i)
#       r=as.numeric(dccmidas_est[[what]][i,i,])
#       r <- cbind(data[c(1:length(r)),1],r)
#       r <- r[,-1]
#       # print(length(r))
#       df=df[c(1:length(r)),]
#       df <- cbind(df, r[,1])
#     }
#     df <- df[-c(1:K_c), ]
#     return(df)
#   }
#   
#   # Other preparations-------------------------------------------------------------------------------------------------------
  t=as.list(label)
  tt=list()
  for (i in 1:n) {
    if (i+1 <= n) {
      for (j in (i+1):n){
        # print(i)
        # print(j)
        tt=append(tt, paste(t[i],"-",t[j]))

      }}}
  tt <- c("Date", tt)

  df_l=d[-1,1]
  df_s=d[-1,1]
  df_v=d[-1,1]
  df_var=d[-1,1]

  r_t <- data[, label]
  r_t <- r_t[complete.cases(r_t),]#去除缺失值
  r_t=r_t[1:(nrow(r_t) - 30), ]
  r_t <- lapply(r_t, function(x) x)

  MV <- lapply(r_t, function(x) mv_into_mat(x, low, K = K, "monthly"))



  # Modeling-------------------------------------------------------------------------------------------------------------------

  dccmidas_est=dcc_fit(r_t,univ_model="gjrGARCH",distribution="norm",corr_model="cDCC")
  dccmidas_est[["corr_coef_mat"]]
  mean=dccmidas_est[["corr_coef_mat"]][1,1]+dccmidas_est[["corr_coef_mat"]][2,1]
  mean
  va=dccmidas_est[["corr_coef_mat"]][1,2]
  vb=dccmidas_est[["corr_coef_mat"]][2,2]
  va
  vb
  cov= -6.171308e-06
  std=(va^2+vb^2+2*cov)^(1/2)
  std
  stat=round(2*(1-stats::pnorm(abs(mean/std))),6)
  stat
  summary.dccmidas(dccmidas_est)
#  
  
#   # Save-------------------------------------------------------------------------------------------------------------------
#   # plot_dccmidas(
#   #   dccmidas_est,
#   #   K_c = K_c,
#   #   vol_col = "black",
#   #   long_run_col = "red",
#   #   cex_axis = 0.75,
#   #   LWD = 2,
#   #   asset_sub = NULL
#   # )
#   
#   
#   df_l=Cor(df_l,"R_t_bar")
#   df_s=Cor(df_s,"R_t")
#   Covar=Cov(df_v,"H_t")
#   Var=Var(df_var,"H_t")
#   names(df_l) <- tt
#   names(df_s) <- tt
#   names(Covar) <- tt
#   names(Var) <- c("Date", label)
#   
#   
#   
#   wb <- createWorkbook()
#   addWorksheet(wb, "Long-Cor")
#   writeData(wb, "Long-Cor", df_l)
#   addWorksheet(wb, "Short-Cor")
#   writeData(wb, "Short-Cor", df_s)
#   addWorksheet(wb, "Covar")
#   writeData(wb, "Covar", Covar)
#   addWorksheet(wb, "Var")
#   writeData(wb, "Var", Var)
#   # saveWorkbook(wb, paste("C:\\Users\\徐浩然\\Desktop\\DCC-data-China\\Varanice&Covariance-",lowname[z],".xlsx",sep = ''), overwrite = TRUE)
#   # saveWorkbook(wb, paste("C:\\Users\\徐浩然\\Desktop\\DCC-data-US\\Varanice&Covariance-",lowname[z],".xlsx",sep = ''), overwrite = TRUE)
#   # saveWorkbook(wb, paste("C:\\Users\\徐浩然\\Desktop\\DCC-data-U\\Varanice&Covariance-",lowname[z],".xlsx",sep = ''), overwrite = TRUE)
#   
#   print('Processing...')
# }



# sink()