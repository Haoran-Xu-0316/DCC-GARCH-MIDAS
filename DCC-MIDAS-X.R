library(dccmidas)
require(xts)
library(readxl)
library(ggplot2)
library(openxlsx)
require(rumidas)

#Hedging Assets----------------------------------------------------------------------------------------------------------
label <- c("沪深300",	"中证500",	"中证1000",	"Green Bond",	"Gold",	"WTI",	"Bitcon","GB")
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
low <- xts(low$CPU, order.by = datess)
# low <- xts(low$EMV, order.by = datess)
# low <- xts(low$TPU, order.by = datess)
# low <- low[-c(1:50), ]
#Other preparations-------------------------------------------------------------------------------------------------------
r_t <- data[, label]
r_t <- r_t[complete.cases(r_t),]#去除缺失值
r_t <- lapply(data[, ], function(x) x)
MV <- lapply(r_t, function(x) mv_into_mat(x, low, K = 24, "monthly"))

K_c<-244
N_c<-22

#Modeling-------------------------------------------------------------------------------------------------------------------
dccmidas_est=dcc_fit(r_t,univ_model="GM_skew",distribution="norm",
                     MV=MV,K=12,corr_model="DCCMIDAS",N_c=N_c,K_c=K_c)

# dccmidas_est=dcc_fit(r_t,univ_model="gjrGARCH",corr_model="aDD",distribution="norm")

# Long Correlation------------------------------------------------------------------------------
Cor=function(df,what){
  for (i in 1:n) {
    if (i+1 <= n) {
      for (j in (i+1):n){
        print(i)
        print(j)
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
  print(count)
  df <- df[-c(1:count-1), ]
  return(df)
}
df_l=Cor(df_l,"R_t_bar")
df_s=Cor(df_s,"R_t")
#Results-------------------------------------------------------------------------------------------------------------------
dccmidas_est
summary.dccmidas(dccmidas_est)
#Plot------------------------------------------------------------------------------
plot_dccmidas(
  dccmidas_est,
  K_c =K_c,
  vol_col = "black",
  long_run_col = "red",
  cex_axis = 0.75,
  LWD = 1.5,
  asset_sub = NULL
)

# save---------------------------------------------------------------------------------------------
wb <- createWorkbook()
addWorksheet(wb, "Long-Cor")
writeData(wb, "Long-Cor", df_l)
addWorksheet(wb, "Short-Cor")
writeData(wb, "Short-Cor", df_s)
saveWorkbook(wb, "C:\\Users\\徐浩然\\Desktop\\DCC data\\DCC.xlsx", overwrite = TRUE)


# write.xlsx(df, "C:\\Users\\徐浩然\\Desktop\\DCC data\\DCC.xlsx", sheetName = "DCC")
