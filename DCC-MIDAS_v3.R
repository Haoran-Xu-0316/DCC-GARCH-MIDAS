#封装过的函数
#输入宏观不确定性，K_c和N_c
#绘图，保存短期和长期相关系数至excel
library(dccmidas)
require(xts)
library(readxl)
library(ggplot2)
library(openxlsx)
require(rumidas)


#Data------------------------------------------------------------------------------------------------------------------
label <- c("CSI 300",	"S&P 500",	"中证1000",	"Gold",	"WTI",	"Bitcoin","Bond")
data <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC-data\\Hedging Assets\\hedging assets.xlsx", sheet = "All-log")
d=data
dates <- as.Date(data$Date, format = "%Y%m%d") 
data <- xts(data[, label], order.by = dates) 
C=c(1,4,5,6,7)
data=data[, C]
label=label[C]
n=length(label)

low <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC-data\\Macro Uncertainty\\macro uncertainty.xlsx", sheet = "All")
datess <- as.Date(low$Date, format = "%Y%m%d") 

# low <- xts(low$EMV, order.by = datess)
# low <- xts(low$USEPU, order.by = datess)
# low <- xts(low$TPU, order.by = datess)
# low <- xts(low$GPR, order.by = datess)
# # low <- xts(low$GVZ, order.by = datess)
# low <- xts(low$OVX, order.by = datess)
# low <- xts(low$VIX, order.by = datess)
# low <- diff(log(low))*100




#Get Cor-------------------------------------------------------------------------------------------------------

diff_macro=function(macro,K,N,lag){
  
  low <- xts(low[macro], order.by = datess)
  low=diff(low)
  low <- low[complete.cases(low),]
  df_l=d[-1,1]
  df_s=d[-1,1]
  
  r_t <- data[, label]
  r_t <- r_t[complete.cases(r_t),]#去除缺失值
  r_t=r_t[1:(nrow(r_t) - 41), ]
  r_t <- head(r_t, n = nrow(r_t) - 0)
  r_t <- lapply(r_t, function(x) x)
  MV <- lapply(r_t, function(x) mv_into_mat(x, low, K = lag, "monthly"))
  
  K_c<-K
  N_c<-N
  
  #Modeling-------------------------------------------------------------------------------------------------------------------
  dccmidas_est=dcc_fit(r_t,univ_model="GM_skew",distribution="norm",
                       MV=MV,K=lag,corr_model="DCCMIDAS",N_c=N_c,K_c=K_c)
  
  # Long Correlation----------------------------------------------------------------------------------------------------------
  Cor=function(df,what){
    for (i in 1:n) {
      if (i+1 <= n) {
        for (j in (i+1):n){
          # print(i)
          # print(j)
          r=as.numeric(dccmidas_est[[what]][i,j,])
          r <- cbind(data[,1],r)
          r <- r[-1,-1]
          df <- cbind(df, r[,1])
        }
      }
    }
    count <- 0
    for (i in 1:nrow(df)) {
      if (df[,2][i] == 0) {
        next
      } else {
        count <- i
        break}
    }
    # print(count)
    df <- df[-c(1:count-1), ]
    return(df)
  }
  df_l=Cor(df_l,"R_t_bar")
  df_s=Cor(df_s,"R_t")
  #Results------------------------------------------------------------------------------------------------------------------
  dccmidas_est
  summary.dccmidas(dccmidas_est)
  #Plot---------------------------------------------------------------------------------------------------------------------
  # plot_dccmidas(
  #   dccmidas_est,
  #   K_c =K_c,
  #   vol_col = "black",
  #   long_run_col = "red",
  #   cex_axis = 0.75,
  #   LWD = 1.5,
  #   asset_sub = NULL)
  
  
  
  # save---------------------------------------------------------------------------------------------------------------------
  file_name <- paste("C:\\Users\\徐浩然\\Desktop\\DCC-data\\DCC","-", macro, ".xlsx", sep = "")
  wb <- createWorkbook()
  addWorksheet(wb, "Long-Cor")
  writeData(wb, "Long-Cor", df_l)
  addWorksheet(wb, "Short-Cor")
  writeData(wb, "Short-Cor", df_s)
  saveWorkbook(wb, file_name, overwrite = TRUE)
  
}

#输入宏观不确定性，K_c和N_c
GPR=diff_macro("GPR",244,22,12)
EPU=diff_macro("EPU",244,22,12)
# CEPU=diff_macro("CEPU",244,22,12)
CPU=diff_macro("CPU",244,22,12)
EMV=diff_macro("EMV",244,22,12)
TPU=diff_macro("TPU",244,22,12)

