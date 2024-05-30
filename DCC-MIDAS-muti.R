library(dccmidas)
require(xts)
library(readxl)
library(ggplot2)
library(openxlsx)
require(rumidas)

label <- c("600519.sh", "600887.sh", "600008.sh", "000002.sz")
n=length(label)
data <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC-GARCH DATA.xlsx", sheet = "data")
dates <- as.Date(data$trade_date, format = "%Y%m%d") 
data <- xts(data[, label], order.by = dates) 

low <- read_excel("C:\\Users\\徐浩然\\Desktop\\China_Mainland_Paper_EPU.xlsx", sheet = "TPU 2000 onwards")
datess <- as.Date(low$date, format = "%Y%m%d") 
low <- xts(low$TPU, order.by = datess) 


r_t <- data[, label]
# r_t <- r_t[complete.cases(r_t),]#去除缺失值
r_t <- lapply(data[, ], function(x) x)

x=r_t

MV <- lapply(r_t, function(x) mv_into_mat(x, low, K = 6, "monthly"))

K_c<-252
N_c<-5

dccmidas_est=dcc_fit(r_t,univ_model="GM_noskew",distribution="norm",
                     MV=MV,K=6,corr_model="DCCMIDAS",N_c=N_c,K_c=K_c)
dccmidas_est
summary.dccmidas(dccmidas_est)
x=dccmidas_est[["R_t"]]
# s=as.numeric(dccmidas_est[["R_t"]][1,2,])
# s <- cbind(r_t_s[,1],s)
# s <- s[-1,-1]
# s <- data.frame(Date = time(s), Value = as.vector(s))
# 
# l=as.numeric(dccmidas_est[["R_t_bar"]][1,2,])
# l <- cbind(r_t_s[,1],l)
# l <- l[-1,-1]
# l <- data.frame(Date = time(l), Value = as.vector(l))
# 
# v1=as.numeric(dccmidas_est[["H_t"]][1,2,])
# v1 <- cbind(r_t_s[,1],v1)
# v1 <- v1[-1,-1]
# v1 <- data.frame(Date = time(v1), Value = as.vector(v1))
# 
# v2=as.numeric(dccmidas_est[["H_t"]][2,1,])
# v2 <- cbind(r_t_s[,1],v2)
# v2 <- v2[-1,-1]
# v2 <- data.frame(Date = time(v2), Value = as.vector(v2))
# write.xlsx(dcr, "C:\\Users\\徐浩然\\Desktop\\Cor-DCC-GARCH.xlsx", sheetName = "Sheet1")
# ------------------------------------------------------------------------------
# 创建一个空数据框，用于存储数据
# dataframe <- data.frame(Date=character(), stringsAsFactors=FALSE)
# for (i in 1:n) {
#   col_name <- paste0("Var", i)
#   dataframe[[col_name]] <- numeric()
# }
# 
# # 使用循环逐个提取每个日期的矩阵数据并添加到数据框中
# for (i in 1:3) {
#   # 获取日期
#   for (j in i+1:3){
#   date <- rownames(x)[i]
#   
#   # 从相关系数矩阵中提取所需数据
#   matrix_data <- as.vector(x[i,j,])
#   
#   # 将日期和相关系数矩阵中的数据合并
#   data_row <- c(date, matrix_data)
#   
#   # 将数据添加到数据框中
#   dataframe <- rbind(dataframe, data_row)
#   }
# }
# 
# # 将数据框中的数值列转换为数值类型
# dataframe[,2:ncol(dataframe)] <- sapply(dataframe[,2:ncol(dataframe)], as.numeric)
# 
# # 按日期排序数据框
# dataframe <- dataframe[order(dataframe$Date),]
# plot(dataframe)
# setwd("C:\\Users\\徐浩然\\Desktop") 
# png(filename = "ROC.png")
# # ------------------------------------------------------------------------------
plot_dccmidas(
  dccmidas_est,
  K_c =K_c,
  vol_col = "black",
  long_run_col = "red",
  cex_axis = 0.75,
  LWD = 1.5,
  asset_sub = NULL
)
# 
# wb <- createWorkbook()
# 
# addWorksheet(wb, "SCor")
# writeData(wb, "SCor", dcr)
# addWorksheet(wb, "LCor")
# writeData(wb, "LCor", dcrb)
# addWorksheet(wb, "v1")
# writeData(wb, "v1", v1)
# addWorksheet(wb, "v2")
# writeData(wb, "v2", v2)

# saveWorkbook(wb, "C:\\Users\\徐浩然\\Desktop\\DCC-GARCH.xlsx", overwrite = TRUE)

# dev.off()

