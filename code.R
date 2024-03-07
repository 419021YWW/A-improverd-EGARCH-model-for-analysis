setwd("D:\\manuscript\\index")
filenames=dir()
library(xlsx)
library(openxlsx)
library(dplyr)
library(readxl)
library(xts)
library(urca)
library(FinTS)
library(forecast)
library(tseries)
library(e1071)
library(TSA)
library(trend)
library(lawstat)
library(randtests)
library(rugarch)
library(vars)
library(frequencyConnectedness)
f=grep('\\.xlsx', filenames, value = TRUE)#正则表达式获取所有xlsx结尾的文件
f
sheet=list()
path=c()
####读取文档
for (i in c(1:36)){
  path[i]=paste0(getwd(),'\\',f[i])
  sheet[[i]]=read_excel(path=path[i])
  names(sheet[[i]])=paste("X",1:9,sep='')}
shee=list()
for(i in c(1:36)){
  shee[[i]]=na.omit(data.frame(as.Date(sheet[[i]]$X1),sheet[[i]]$X7))
  names(shee[[i]])=c("date","closeprice")
}
n=c()
she=list()
he=list()
it=list()
all_results <- data.frame()
all_results1 <- data.frame()
all_results2 <- data.frame()
all_results3 <- data.frame()
all_results4 <- data.frame()
all_results5 <- data.frame()
#计算股票的对数收益率
for(m in 1:36){
  n[m]=nrow(shee[[m]])
  she[[m]]=log(((shee[[m]]$closeprice[2:n[m]])/(shee[[m]]$closeprice[1:n[m]-1])))*100
  he[[m]]=shee[[m]][-1,]$date
  it[[m]]=data.frame(she[[m]],he[[m]])
  all_results=rbind(all_results,summary(she[[m]]))
  all_results1=rbind(all_results1,sd(she[[m]]))
  all_results2=rbind(all_results2,length(she[[m]]))
  all_results3=rbind(all_results3,adf.test(she[[m]])$statistic)
  all_results4=rbind(all_results4,skewness(she[[m]]))
  all_results5=rbind(all_results5,kurtosis(she[[m]]))
}
#write.csv(all_results,file="summary.csv",row.names = FALSE)
#write.csv(all_results1,file="sd.csv",row.names = FALSE)
#write.csv(all_results2,file="n.csv",row.names = FALSE)
#write.csv(all_results3,file="adf.csv",row.names = FALSE)
#write.csv(all_results4,file="skewness.csv",row.names = FALSE)
#write.csv(all_results5,file="kurtosis.csv",row.names = FALSE)
newsheet=list()
for(n in 1:36){
  newsheet[[n]]=xts(it[[n]]$she..m..,order.by=as.Date(it[[n]]$he..m..))
}
#arima滞后阶数检验，arima检验
for(p in 1:36){
  print(arima(newsheet[[p]]))
} 
auto.arima(newsheet[[30]])
#提取滞后阶数
lags=matrix(c(1,2),c(2,3),c(2,3),c(0,2),c(1,2),c(1,1),c(5,2),c(0,1),
            c(0,0),c(3,2),c(5,2),c(2,1),c(1,1),c(5,1),c(3,2),c(1,2),
            c(2,3),c(0,0),c(5,4),c(0,2),c(0,0),c(4,4),c(3,2),c(0,0),
            c(1,1),c(5,0),c(2,1),c(5,1),c(5,2),c(0,0),c(2,2),c(0,2),
            c(0,2),c(0,1),c(3,3),c(3,5))
lags1=c(1,2,2,0,1,1,5,0,0,3,5,2,1,5,3,1,2,0,5,0,0,4,3,0,1,5,2,5,5,0,2,0,0,0,3,3)
#length(lags2)
lags2=c(2,3,3,2,2,1,2,1,0,2,2,1,1,1,2,2,3,0,4,2,0,4,2,0,1,0,1,1,2,0,2,2,2,1,3,5)
####提取中美EPU指数
setwd("D:\\manuscript\\EPU\\daily EPU")
EPU=read.csv("epu_daily_18_june_2022_updated.csv")
EPU1=na.omit(EPU)
nrow(EPU1)
#计算偏度和峰度
skewness(EPU1$USAEPU_Daily)
skewness(EPU1$CNEPU_Daily)
kurtosis(EPU1$USAEPU_Daily)
kurtosis(EPU1$CNEPU_Daily)
#转化为时间序列数据
epu_us_g=xts(EPU1$USAEPU_Daily,order.by=as.Date(EPU1$date))
epu_g=xts(EPU1$CNEPU_Daily,order.by=as.Date(EPU1$date))
epulog=log(epu_g)
epuusg=log(epu_us_g)
#epu_g=epu_g[!is.infinite(rowSums(epu_g)),]
summary(epu_us_g)
summary(epu_g)
adf.test(epu_us_g)
adf.test(epu_g)
#一带一路政策提出前
epu_us_g1=xts(EPU1$USAEPU_Daily,order.by=as.Date(EPU1$date))['2000-01-01/2013-09-01']
epu_g1=xts(EPU1$CNEPU_Daily,order.by=as.Date(EPU1$date))['2000-01-01/2013-09-01']
epulog1=log(epu_g1)
epuusg1=log(epu_us_g1)
#一带一路政策提出后2017年5月1日-2022年6月
epu_us_g2=xts(EPU1$USAEPU_Daily,order.by=as.Date(EPU1$date))['2017-05-01/2022-05-01']
epu_g2=xts(EPU1$CNEPU_Daily,order.by=as.Date(EPU1$date))['2017-05-01/2022-05-01']
epulog2=log(epu_g2)
epuusg2=log(epu_us_g2)
epu_us_g2_=xts(EPU1$USAEPU_Daily,order.by=as.Date(EPU1$date))['2017-05-01/2019-05-01']
epu_g2_=xts(EPU1$CNEPU_Daily,order.by=as.Date(EPU1$date))['2017-05-01/2019-05-01']
epulog2_=log(epu_g2_)
epuusg2_=log(epu_us_g2_)
#2008年金融危机
epu_us_g3=xts(EPU1$USAEPU_Daily,order.by=as.Date(EPU1$date))['2008-01-01/2008-12-31']
epu_g3=xts(EPU1$CNEPU_Daily,order.by=as.Date(EPU1$date))['2008-01-01/2008-12-31']
epulog3=log(epu_g3)
epuusg3=log(epu_us_g3)
#欧债危机2011年1月至2013年12月
epu_us_g4=xts(EPU1$USAEPU_Daily,order.by=as.Date(EPU1$date))['2011-01-01/2013-12-31']
epu_g4=xts(EPU1$CNEPU_Daily,order.by=as.Date(EPU1$date))['2011-01-01/2013-12-31']
epulog4=log(epu_g4)
epuusg4=log(epu_us_g4)
#中国EPU对一带一路沿线国家整体的影响
mergeindex=list()
mer=list()
dat=list()
for(q in 1:36){
  mergeindex[[q]]=na.omit(merge(newsheet[[q]],epulog,epuusg))
  dat[[q]]=index(mergeindex[[q]])
  mer[[q]]=cbind(data.frame(dat[[q]],mergeindex[[q]]))
}
country.fit=list()
country.spec=list()
start1=Sys.time()
for(k in 1:36){
  country.spec[[k]]=ugarchspec(variance.model=list(model="eGARCH",garchOrder=c(1,1),external.regressors=mergeindex[[k]]$epulog),mean.model=list(armaOrder=c(0,0),external.regressors=mergeindex[[k]]$epulog),distribution.model="norm")
  country.fit[[k]]=ugarchfit(spec=country.spec[[k]],data=mergeindex[[k]]$newsheet..q..,solver = "solnp")
  print(country.fit[[k]])
}
#一带一路前中国EPU对一带一路沿线国家的影响
mergeindex1=list()
mer1=list()
dat1=list()
for(q in 1:36){
  mergeindex1[[q]]=na.omit(merge(newsheet[[q]],epulog1,epuusg1))
  dat1[[q]]=index(mergeindex1[[q]])
  mer1[[q]]=cbind(data.frame(dat1[[q]],mergeindex1[[q]]))
}
country.fit1=list()
country.spec1=list()
for(k in 1:36){
  country.spec1[[k]]=ugarchspec(variance.model=list(model="eGARCH",garchOrder=c(1,1),external.regressors=mergeindex1[[k]]$epulog),mean.model=list(armaOrder=c(0,0),external.regressors=mergeindex1[[k]]$epulog),distribution.model="norm")
  country.fit1[[k]]=ugarchfit(spec=country.spec1[[k]],data=mergeindex1[[k]]$newsheet..q..,solver = "solnp")
}
#一带一路后中国EPU对一带一路沿线国家的影响
mergeindex2=list()
mer2=list()
dat2=list()
for(q in 1:36){
  mergeindex2[[q]]=na.omit(merge(newsheet[[q]],epulog2,epuusg2))
  dat2[[q]]=index(mergeindex2[[q]])
  mer2[[q]]=cbind(data.frame(dat2[[q]],mergeindex2[[q]]))
}
country.fit2=list()
country.spec2=list()
for(k in 1:36){
  country.spec2[[k]]=ugarchspec(variance.model=list(model="eGARCH",garchOrder=c(1,1),external.regressors=mergeindex2[[k]]$epulog),mean.model=list(armaOrder=c(0,0),external.regressors=mergeindex2[[k]]$epulog),distribution.model="norm")
  country.fit2[[k]]=ugarchfit(spec=country.spec2[[k]],data=mergeindex2[[k]]$newsheet..q..,solver = "solnp")
  print(country.fit2[[k]])
}
#中美两国EPU对一带一路沿线国家的影响金融危机、欧债危机
#金融危机
mergeindex3=list()
mer3=list()
dat3=list()
for(q in 1:36){
  mergeindex3[[q]]=na.omit(merge(newsheet[[q]],epulog3,epuusg3))
  dat3[[q]]=index(mergeindex3[[q]])
  mer3[[q]]=cbind(data.frame(dat3[[q]],mergeindex3[[q]]))
}
logexter3=list()
country.fit3=list()
country.spec3=list()
for (k in 1:36) {
  tryCatch({
    mer3[[k]]$epulog = ifelse(mer3[[k]]$epulog != -Inf, mer3[[k]]$epulog, median(mer3[[k]]$epulog))
    mer3[[k]]$epuusg = ifelse(mer3[[k]]$epuusg != -Inf, mer3[[k]]$epuusg, median(mer3[[k]]$epuusg))
    logexter3[[k]] = cbind(data.frame(mer3[[k]])$epulog, data.frame(mer3[[k]])$epuusg)
    country.spec3[[k]] = ugarchspec(
      variance.model = list(model = "eGARCH", garchOrder = c(1, 1), external.regressors = logexter3[[k]]),
      mean.model = list(armaOrder = c(lags1[[k]], lags2[[k]]), external.regressors = logexter3[[k]]),
      distribution.model = "norm"
    )
    country.fit3[[k]] = ugarchfit(spec = country.spec3[[k]], data = mer3[[k]]$newsheet..q.., solver = "gosolnp")
  }, error = function(e) {
    cat("Error in iteration", k, ": ", conditionMessage(e), "\n")
  })
}
#欧债危机
mergeindex4=list()
mer4=list()
dat4=list()
for(q in 1:36){
  mergeindex4[[q]]=na.omit(merge(newsheet[[q]],epulog4,epuusg4))
  dat4[[q]]=index(mergeindex4[[q]])
  mer4[[q]]=cbind(data.frame(dat4[[q]],mergeindex4[[q]]))
}
logexter4=list()
country.fit4=list()
country.spec4=list()
for(k in 1:36){
  mer4[[k]]$epulog=ifelse(mer4[[k]]$epulog!=-Inf,mer4[[k]]$epulog,median(mer4[[k]]$epulog))
  mer4[[k]]$epuusg=ifelse(mer4[[k]]$epuusg!=-Inf,mer4[[k]]$epuusg,median(mer4[[k]]$epuusg))
  logexter4[[k]]=cbind(data.frame(mer4[[k]])$epulog,data.frame(mer4[[k]])$epuusg)
  country.spec4[[k]]=ugarchspec(variance.model=list(model="eGARCH",garchOrder=c(1,1),
                                                   external.regressors=logexter4[[k]]),mean.model=list(armaOrder=c(lags1[[k]],lags2[[k]]),external.regressors=logexter4[[k]]),distribution.model="norm")
  country.fit4[[k]]=ugarchfit(spec=country.spec4[[k]],data=mer4[[k]]$newsheet..q..,solver = "gosolnp")
}
#######################################
###中国美国epu 加入疫情虚拟变量、乘法方式加入
epu_us_g6=xts(EPU1$USAEPU_Daily,order.by=as.Date(EPU1$date))['/2022-06-01']
epu_g6=xts(EPU1$CNEPU_Daily,order.by=as.Date(EPU1$date))['/2022-06-01']
epulog6=log(epu_g6)
epuusg6=log(epu_us_g6)
mergeindex6=list()
for(q in 1:36){
  mergeindex6[[q]]=na.omit(merge(newsheet[[q]],epulog6,epuusg6))
}
epu_us=list()
epulist=list()
coviddata=list()
n=c()
m=c()
dummy=list()
dummydata=list()
dummycountry.spec=list()
dummycountry.fit=list()
for(d in 1:36){
  mergeindex6[[d]]$epulog6=ifelse(mergeindex6[[d]]$epulog6!=-Inf,mergeindex6[[d]]$epulog6,median(mergeindex6[[d]]$epulog6))
  mergeindex6[[d]]$epuusg6=ifelse(mergeindex6[[d]]$epuusg6!=-Inf,mergeindex6[[d]]$epuusg6,median(mergeindex6[[d]]$epuusg6))
  epulist[[d]]=cbind(data.frame(mergeindex6[[d]])$epulog6)
  epu_us[[d]]=cbind(data.frame(mergeindex6[[d]])$epuusg6)
  coviddata[[d]]=mergeindex6[[d]]["2020-01-01/"] 
  n[d]=nrow(coviddata[[d]])
  m[d]=nrow(mergeindex6[[d]])
  dummy[[d]]=as.matrix(c(rep(0,m[d]-n[d]),rep(1,n[d])))
  dummydata[[d]]=cbind(epulist[[d]],dummy[[d]],epu_us[[d]],epulist[[d]]*dummy[[d]],epu_us[[d]]*dummy[[d]])
  dummycountry.spec[[d]]=ugarchspec(variance.model=list(model="eGARCH",garchOrder=c(1,1),
                                                        external.regressors=dummydata[[d]]),mean.model=list(armaOrder=c(lags1[[d]],lags2[[d]]),external.regressors=dummydata[[d]]),distribution.model="norm")
  dummycountry.fit[[d]]=ugarchfit(spec=dummycountry.spec[[d]],data=mergeindex6[[d]]$newsheet..q..,solver = "gosolnp")
}     
###############网络分析##########################
data=read_excel("D:\\manuscript\\data.xlsx")
df <- na.omit(data[,2:38])
df <- df %>%
  mutate_at(vars(3:37), ~log(. / lag(.))) %>%
  filter_all(all_vars(!is.infinite(.))) %>%
  filter_all(all_vars(complete.cases(.)))
est <- VAR(df, p = 2, type = "const")
Table_4=spilloverDY12(est,n.ahead = 12, no.corr = FALSE)
Table_4$tables
write.csv(Table_4$tables,file="D:\\spillover1.csv",row.names =TRUE)
sp1pairwise<-pairwise(Table_4)
sp1pairwise
sink('sp1pairwise.txt') 
print(sp1pairwise)
bounds <- c(pi+0.00001, pi/1,pi/3,pi/6, 0)
Tabel_5<- spilloverBK12(est,n.ahead = 100, no.corr = F, partition = bounds)
Tabel_5
spilloverBK12Anet<-net(Tabel_5)
spilloverBK12Anet
write.csv(Tabel_5$tables,file="D:\\SpilloverBK.csv",row.names =TRUE)
spilloverBK12Apairwise<-pairwise(Tabel_5) 
spilloverBK12Apairwise
plotOverall(Tabel_5)
write.csv(sp1pairwise,file="D:\\1.csv",row.names =TRUE)
library(igraph)
nodes=read.xlsx("D:\\manuscript\\nodes.xlsx")
links=read.xlsx("D:\\manuscript\\edges.xlsx")
net=graph_from_data_frame(d=links,vertices=nodes,directed=T)
library(scales)
edge_weights <- links$weight
cut.off=0 
net.sp=delete_edges(net, E(net)[weight<cut.off]) 
#edge_weights_scaled <- rescale(edge_weights)  
E(net.sp)$width=E(net.sp)$weight*15
E(net.sp)$color <- "black"
V(net.sp)$color <- ifelse(V(net.sp)$name == "CN", "#F39B7FFF", "#91D1C2FF")
#V(net.sp)$color <- ifelse(V(net.sp)$name == "USA", "#E64B35FF", "#4DBBD5FF")
plot(net.sp,
     edge.arrow.size=0.5,
     vertex.label=V(net)$name, 
     vertex.label.color="black",
     layout=layout.davidson.harel,
     vertex.label.cex=.6)
