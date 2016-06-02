library(readxl)
library(plyr)
library(dplyr)
library(reshape2)
library(forecast)
library(corrplot)
library(car)
library(coefplot)

########
dr <- read_excel('C:/Users/aarschle1/Google Drive/Optimedia/T-Mobile/Projects/Weekly_Reporting/TMO Q1 Master Pivot 1.1.16-1.24.16_working.xlsm', 'data')

dat <- dr[,c(2,5,9,31:34,53)]

dat <- melt(dat, id.vars = c('Date', 'Campaign', 'Site'))
dat <- dcast(dat, formula = Campaign + Site + Date ~ variable, fun.aggregate = sum)
########

reporting_date <- as.Date(max(dr$Date))
end_date <- as.Date('2016-03-31')

ahead <- end_date - reporting_date

dat <- dat[order(dat$Campaign, dat$Site, dat$Date),]

campaigns.split <- split(dat, dat$Campaign)

spend.forecast <- lapply(campaigns.split, function(x) {
  y <- split(x, x$Site)
  z <- lapply(y, function(t) {
    a <- ts(t[,4:8], frequency = 7)
    e <- tslm(Impressions ~ NTC.Media.Cost, data = a)
    b <- auto.arima(a[,1])
    c <- forecast(b, h = ahead)
    d <- data.frame(c)
    names(d) <- c("NTC Media Cost", "Lo80", "Hi80", "Lo95", "Hi95")
    f <- forecast(e, newdata = d)
    g <- data.frame(f)
    names(g) <- c("Impressions", "Lo80", "Hi80", "Lo95", "Hi95")
    h <- data.frame(cbind(d, g))
    return(h)
  })
})

forecast.df <- as.data.frame(spend.forecast)
forecast.df <- forecast.df[,-grep('Hi|Lo', colnames(forecast.df))]

save.xlsx('C:/Users/aarschle1/Desktop/pacing_forecast.xlsx', forecast.df)
