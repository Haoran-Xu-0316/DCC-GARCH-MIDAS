#为了获取单个Double Asymmetry Garch Midas 的短期波动和长期波动
#不一定会使用

library(rumidas)
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
x=data
x <- x[complete.cases(x),]
x=x[1:(nrow(x) - 41), ]

dates <- as.Date(data$Date, format = "%Y%m%d") 

data <- xts(data[, label], order.by = dates) 
C=c(1,4,5,6,7)
# C=c(1,2,3,4,5)
# C=c(1,8,4,5,6,7)
data=data[, C]*1
label=label[C]
n=length(label)

#Macro Uncertainty--------------------------------------------------------------------------------------------------------
low <- read_excel("C:\\Users\\徐浩然\\Desktop\\DCC-data\\Macro Uncertainty\\macro uncertainty.xlsx", sheet = "All")

low <- low[complete.cases(low),]
datess <- as.Date(low$Date, format = "%Y%m%d")

#Other preparations-------------------------------------------------------------------------------------------------------

Vol=function(asset,macro,K){
  
  
  r_t = data[,asset]
  # r_t <- r_t[complete.cases(r_t),]
  r_t=r_t[1:(nrow(r_t) - 41), ]
  r_t <- head(r_t, n = nrow(r_t) - 0)
  low <- xts(low[,macro], order.by = datess)
  low=diff(low)
  
  mv_m<-mv_into_mat(r_t,low,K=K,"monthly")
  fit<-ugmfit(model="DAGM",skew="YES",distribution="norm",r_t,mv_m,K=K)
  
  summary.rumidas(fit)
  print(fit["loglik"])
  print(fit["inf_criteria"])


  est=unlist(fit["rob_coef_mat"])

  alpha=est["rob_coef_mat.Estimate1"]
  gamma=est["rob_coef_mat.Estimate2"]
  beta=est["rob_coef_mat.Estimate3"]
  m=est["rob_coef_mat.Estimate4"]
  theta_pos=est["rob_coef_mat.Estimate5"]
  w2_pos=est["rob_coef_mat.Estimate6"]
  theta_neg=est["rob_coef_mat.Estimate7"]
  w2_neg=est["rob_coef_mat.Estimate8"]

  param<-c(alpha=alpha,gamma=gamma,beta=beta,m=m,theta_pos=theta_pos,w2_pos=w2_pos,theta_neg=theta_neg,w2_neg=w2_neg)

  SV=DAGM_cond_vol(param,r_t,mv_m,K)
  LV=DAGM_long_run_vol(param, r_t, mv_m, K)

  SVLV <- list(SV,LV)
  return(SVLV)
}
SCPU=Vol("CSI 300","CPU",11)
SEMU=Vol("CSI 300","EMV",29)
SEPU=Vol("CSI 300","EPU",19)
STPU=Vol("CSI 300","TPU",23)
SGPR=Vol("CSI 300","GPR",30)

GCPU=Vol("Gold","CPU",13)
GEMU=Vol("Gold","EMV",19)
GEPU=Vol("Gold","EPU",8)
GTPU=Vol("Gold","TPU",19)
GGPR=Vol("Gold","GPR",13)

OCPU=Vol("WTI","CPU",27)
OEMU=Vol("WTI","EMV",29)
OEPU=Vol("WTI","EPU",17)
OTPU=Vol("WTI","TPU",21)
OGPR=Vol("WTI","GPR",26)

BCPU=Vol("Bitcoin","CPU",14)
BEMU=Vol("Bitcoin","EMV",27)
BEPU=Vol("Bitcoin","EPU",16)
BTPU=Vol("Bitcoin","TPU",12)
BGPR=Vol("Bitcoin","GPR",7)

BDCPU=Vol("Bond","CPU",30)
BDEMU=Vol("Bond","EMV",18)
BDEPU=Vol("Bond","EPU",8)
BDTPU=Vol("Bond","TPU",22)
BDGPR=Vol("Bond","GPR",12)


SCPUS=SCPU[[1]]
SCPUL=SCPU[[2]]
SEMUS=SEMU[[1]]
SEMUL=SEMU[[2]]
SEPUS=SEPU[[1]]
SEPUL=SEPU[[2]]
STPUS=STPU[[1]]
STPUL=STPU[[2]]
SGPRS=SGPR[[1]]
SGPRL=SCPU[[2]]

GCPUS=GCPU[[1]]
GCPUL=GCPU[[2]]
GEMUS=GEMU[[1]]
GEMUL=GEMU[[2]]
GEPUS=GEPU[[1]]
GEPUL=GEPU[[2]]
GTPUS=GTPU[[1]]
GTPUL=GTPU[[2]]
GGPRS=GGPR[[1]]
GGPRL=GCPU[[2]]

OCPUS=OCPU[[1]]
OCPUL=OCPU[[2]]
OEMUS=OEMU[[1]]
OEMUL=OEMU[[2]]
OEPUS=OEPU[[1]]
OEPUL=OEPU[[2]]
OTPUS=OTPU[[1]]
OTPUL=OTPU[[2]]
OGPRS=OGPR[[1]]
OGPRL=OCPU[[2]]

BCPUS=BCPU[[1]]
BCPUL=BCPU[[2]]
BEMUS=BEMU[[1]]
BEMUL=BEMU[[2]]
BEPUS=BEPU[[1]]
BEPUL=BEPU[[2]]
BTPUS=BTPU[[1]]
BTPUL=BTPU[[2]]
BGPRS=BGPR[[1]]
BGPRL=BCPU[[2]]

BDCPUS=BDCPU[[1]]
BDCPUL=BDCPU[[2]]
BDEMUS=BDEMU[[1]]
BDEMUL=BDEMU[[2]]
BDEPUS=BDEPU[[1]]
BDEPUL=BDEPU[[2]]
BDTPUS=BDTPU[[1]]
BDTPUL=BDTPU[[2]]
BDGPRS=BDGPR[[1]]
BDGPRL=BDCPU[[2]]


stock=cbind(x[,1],
            SCPUS,
            SCPUL,
            SEMUS,
            SEMUL,
            SEPUS,
            SEPUL,
            STPUS,
            STPUL,
            SGPRS,
            SGPRL)


Gold=cbind(x[,1],
           GCPUS,
           GCPUL,
           GEMUS,
           GEMUL,
           GEPUS,
           GEPUL,
           GTPUS,
           GTPUL,
           GGPRS,
           GGPRL)


oil=cbind(x[,1],
          OCPUS,
          OCPUL,
          OEMUS,
          OEMUL,
          OEPUS,
          OEPUL,
          OTPUS,
          OTPUL,
          OGPRS,
          OGPRL)


btc=cbind(x[,1],
          BCPUS,
          BCPUL,
          BEMUS,
          BEMUL,
          BEPUS,
          BEPUL,
          BTPUS,
          BTPUL,
          BGPRS,
          BGPRL)


bond=cbind(x[,1],
           BDCPUS,
           BDCPUL,
           BDEMUS,
           BDEMUL,
           BDEPUS,
           BDEPUL,
           BDTPUS,
           BDTPUL,
           BDGPRS,
           BDGPRL)

# plot(S)
# lines(L, col = "red",lwd = 2)




wb <- createWorkbook()
addWorksheet(wb, "stock")
writeData(wb, "stock", stock)
addWorksheet(wb, "Gold")
writeData(wb, "Gold", Gold)
addWorksheet(wb, "oil")
writeData(wb, "oil", oil)
addWorksheet(wb, "btc")
writeData(wb, "btc", btc)
addWorksheet(wb, "bond")
writeData(wb, "bond", bond)
saveWorkbook(wb, "C:\\Users\\徐浩然\\Desktop\\DCC-data\\DiffVolatility.xlsx", overwrite = TRUE)






