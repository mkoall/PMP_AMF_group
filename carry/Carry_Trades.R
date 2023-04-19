library(readxl)
library(xts)
library(tidyverse)
library(dplyr)
library(broom)
library(ggpubr)
library(moments)
library(scales)
library(data.table)
library(texreg)

setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
remove(list = ls())
gc()

# ticker <- read_excel("FX_Data_Datastream.xlsx", sheet = "ticker", col_names = F)

SPTS <- read_excel("SpotsForwards.xlsx", sheet = "SPOTS")
FWDS <- read_excel("SpotsForwards.xlsx", sheet = "FWDS")


tickers <- c("Date", "NGN", "BRL", "ARS", "EGP", "ZAR", "KES", "GHS", "INR", "CNY", "IDR")
spts_tickers <- c("Date", paste(tickers[-1], "Spt"))
fwds_tickers <- c("Date", paste(tickers[-1], "Fwd"))

colnames(SPTS) <- spts_tickers
colnames(FWDS) <- fwds_tickers

SPTS <- xts(SPTS[,-1], order.by = SPTS$Date)
FWDS <- xts(FWDS[,-1], order.by = FWDS$Date)

# colnames(ticker) <- c("ticker", "spotuk", "fwduk", "fwdus", "spotus", "EURstart")
# spotus <- read_excel("FX_Data_Datastream.xlsx", sheet = "spots US", skip = 1, na = "NA", guess_max = 10000)
# fwdus <- read_excel("FX_Data_Datastream.xlsx", sheet = "forwards US", skip = 1, na = "NA", guess_max = 10000)

# spotus_xts <- merge(xts(spotus[,2:16], order.by = spotus$Code...1),
#                     xts(spotus[1:length(na.omit(spotus$Code...17)),18:40], order.by = na.omit(spotus$Code...17)),
#                     xts(spotus[1:length(na.omit(spotus$Code...41)),42:49], order.by = na.omit(spotus$Code...41)))
# 
# colnames(spotus_xts) <- ticker$ticker[match(c(paste0(substr(colnames(spotus_xts)[1:10], 1, 6), "$"),
#                                               colnames(spotus_xts)[11],
#                                               paste0(substr(colnames(spotus_xts)[12:46], 1, 6), "$")), ticker$spotus)]
# 
# 
# fwdus_xts <- merge(xts(fwdus[,2:16], order.by = fwdus$Code...1),
#                    xts(fwdus[1:length(na.omit(fwdus$Code...17)),18:41], order.by = na.omit(fwdus$Code...17)),
#                    xts(fwdus[1:length(na.omit(fwdus$Code...42)),43:50], order.by = na.omit(fwdus$Code...42)))
# 
# colnames(fwdus_xts) <- ticker$ticker[match(colnames(fwdus_xts), ticker$fwdus)]


# fwdus_xts[, c("AUD", "EUR", "GBP", "NZD", "IEP")] <- 1 / fwdus_xts[, c("AUD", "EUR", "GBP", "NZD", "IEP")]
# 
# cols <- c("BRL", "CZK", "INR", "MXN", "RUB", "PLN", "ZAR", "PHP", "THB", "TRY")
# start <- "2004-03/"


SPTS <- do.call(merge, lapply(SPTS, to.monthly, OHLC = F))
FWDS <- do.call(merge, lapply(FWDS, to.monthly, OHLC = F))

portfolio <- lag.xts(FWDS, 1) / SPTS - 1
forward_discount <- FWDS / SPTS - 1

colnames(portfolio) <- tickers[-1]
colnames(forward_discount) <- tickers[-1]


ranks <- forward_discount

portfolio_1 <- c()
portfolio_2 <- c()
portfolio_3 <- c()
portfolio_4 <- c()
portfolio_5 <- c()

for(i in 1:(nrow(forward_discount)-1)){
  ranks[i,] <- rank(as.numeric(forward_discount[i,]), ties.method = "average")
  portfolio_1[i+1] <- mean(as.numeric(portfolio[i+1, as.vector(ranks[i,]) < 3]))
  portfolio_2[i+1] <- mean(as.numeric(portfolio[i+1, as.vector(ranks[i,]) < 5 & 2 < as.vector(ranks[i,])]))
  portfolio_3[i+1] <- mean(as.numeric(portfolio[i+1, as.vector(ranks[i,]) < 7 & 4 < as.vector(ranks[i,])]))
  portfolio_4[i+1] <- mean(as.numeric(portfolio[i+1, as.vector(ranks[i,]) < 9 & 6 < as.vector(ranks[i,])]))
  portfolio_5[i+1] <- mean(as.numeric(portfolio[i+1, as.vector(ranks[i,]) > 8]))
}

portfolio1 <- xts(portfolio_1, order.by=index(portfolio))
portfolio2 <- xts(portfolio_2, order.by=index(portfolio))
portfolio3 <- xts(portfolio_3, order.by=index(portfolio))
portfolio4 <- xts(portfolio_4, order.by=index(portfolio))
portfolio5 <- xts(portfolio_5, order.by=index(portfolio))


HML <- (portfolio5 - portfolio1) * 100 * 12
DOLL <- (portfolio$NGN + portfolio$BRL + portfolio$ARS + portfolio$EGP + portfolio$ZAR + portfolio$KES + portfolio$GHS + portfolio$INR + portfolio$CNY + portfolio$IDR) / 10 * 100 * 12
# DOLL <- (portfolio$BRL + portfolio$CZK + portfolio$INR + portfolio$MXN + portfolio$RUB + portfolio$PLN + portfolio$ZAR + portfolio$PHP + portfolio$THB + portfolio$TRY) / 10 * 100 * 12

regression_1 <- lm(portfolio1 * 100 * 12 ~ DOLL + HML)
regression_2 <- lm(portfolio2 * 100 * 12 ~ DOLL + HML)
regression_3 <- lm(portfolio3 * 100 * 12 ~ DOLL + HML)
regression_4 <- lm(portfolio4 * 100 * 12 ~ DOLL + HML)
regression_5 <- lm(portfolio5 * 100 * 12 ~ DOLL + HML)


means <- data.frame(p1 = c(mean(portfolio1, na.rm=TRUE) * 12),
                    p2 = c(mean(portfolio2, na.rm=TRUE) * 12),
                    p3 = c(mean(portfolio3, na.rm=TRUE) * 12),
                    p4 = c(mean(portfolio4, na.rm=TRUE) * 12),
                    p5 = c(mean(portfolio5, na.rm=TRUE) * 12))

std <- data.frame(p1 = c(sd(portfolio1, na.rm=TRUE) * sqrt(12)),
                  p2 = c(sd(portfolio2, na.rm=TRUE) * sqrt(12)),
                  p3 = c(sd(portfolio3, na.rm=TRUE) * sqrt(12)),
                  p4 = c(sd(portfolio4, na.rm=TRUE) * sqrt(12)),
                  p5 = c(sd(portfolio5, na.rm=TRUE) * sqrt(12))) 


skewness <-  data.frame(p1 = skewness(portfolio1, na.rm=TRUE),
                        p2 = skewness(portfolio2, na.rm=TRUE),
                        p3 = skewness(portfolio3, na.rm=TRUE),
                        p4 = skewness(portfolio4, na.rm=TRUE),
                        p5 = skewness(portfolio5, na.rm=TRUE))

kurtosis <- data.frame(p1 = kurtosis(portfolio1, na.rm=TRUE),
                       p2 = kurtosis(portfolio2, na.rm=TRUE),
                       p3 = kurtosis(portfolio3, na.rm=TRUE),
                       p4 = kurtosis(portfolio4, na.rm=TRUE),
                       p5 = kurtosis(portfolio5, na.rm=TRUE))

SR <- data.frame(p1 = means$p1 / std$p1,
                 p2 = means$p2 / std$p2,
                 p3 = means$p3 / std$p3,
                 p4 = means$p4 / std$p4,
                 p5 = means$p5 / std$p5)

summary <- data.frame(Portfolio_1 = c(scales::percent(c(means$p1, std$p1), accuracy = 0.01), round(c(skewness$p1, kurtosis$p1, SR$p1),2)),
                      Portfolio_2 = c(scales::percent(c(means$p2, std$p2), accuracy = 0.01), round(c(skewness$p2, kurtosis$p2, SR$p2),2)),
                      Portfolio_3 = c(scales::percent(c(means$p3, std$p3), accuracy = 0.01), round(c(skewness$p3, kurtosis$p3, SR$p3),2)),
                      Portfolio_4 = c(scales::percent(c(means$p4, std$p4), accuracy = 0.01), round(c(skewness$p4, kurtosis$p4, SR$p4),2)),
                      Portfolio_5 = c(scales::percent(c(means$p5, std$p5), accuracy = 0.01), round(c(skewness$p5, kurtosis$p5, SR$p5),2)))



rownames(summary) <- c("Mean", "Std", "Skewness", "Kurtosis", "SR")

regressions <- list(Portfolio_1 =regression_1,
                    Portfolio_2 =regression_2,
                    Portfolio_3 =regression_3,
                    Portfolio_4 =regression_4,
                    Portfolio_5 =regression_5)

htmlreg(regressions, file = "table.doc")

#spot change calculations

spotchange <- (lag.xts(SPTS, 1) / SPTS) - 1

spotchange_1 <- c()
spotchange_2 <- c()
spotchange_3 <- c()
spotchange_4 <- c()
spotchange_5 <- c()

for(i in 1:(nrow(forward_discount)-1)){
  ranks[i,] <- rank(as.numeric(forward_discount[i,]), ties.method = "average")
  spotchange_1[i+1] <- mean(as.numeric(spotchange[i+1, as.vector(ranks[i,]) < 3]))
  spotchange_2[i+1] <- mean(as.numeric(spotchange[i+1, as.vector(ranks[i,]) < 5 & 2 < as.vector(ranks[i,])]))
  spotchange_3[i+1] <- mean(as.numeric(spotchange[i+1, as.vector(ranks[i,]) < 7 & 4 < as.vector(ranks[i,])]))
  spotchange_4[i+1] <- mean(as.numeric(spotchange[i+1, as.vector(ranks[i,]) < 9 & 6 < as.vector(ranks[i,])]))
  spotchange_5[i+1] <- mean(as.numeric(spotchange[i+1, as.vector(ranks[i,]) > 8]))
}

spotchange1 <- xts(spotchange_1, order.by=index(spotchange))
spotchange2 <- xts(spotchange_2, order.by=index(spotchange))
spotchange3 <- xts(spotchange_3, order.by=index(spotchange))
spotchange4 <- xts(spotchange_4, order.by=index(spotchange))
spotchange5 <- xts(spotchange_5, order.by=index(spotchange))

means_sp <- data.frame(p1 = c(mean(spotchange1, na.rm=TRUE) * 12),
                       p2 = c(mean(spotchange2, na.rm=TRUE) * 12),
                       p3 = c(mean(spotchange3, na.rm=TRUE) * 12),
                       p4 = c(mean(spotchange4, na.rm=TRUE) * 12),
                       p5 = c(mean(spotchange5, na.rm=TRUE) * 12))

std_sp <- data.frame(p1 = c(sd(spotchange1, na.rm=TRUE) * sqrt(12)),
                     p2 = c(sd(spotchange2, na.rm=TRUE) * sqrt(12)),
                     p3 = c(sd(spotchange3, na.rm=TRUE) * sqrt(12)),
                     p4 = c(sd(spotchange4, na.rm=TRUE) * sqrt(12)),
                     p5 = c(sd(spotchange5, na.rm=TRUE) * sqrt(12))) 


summary_sp_ch <- data.frame(Portfolio_1 = c(scales::percent(c(means_sp$p1, std_sp$p1), accuracy = 0.01)),
                            Portfolio_2 = c(scales::percent(c(means_sp$p2, std_sp$p2), accuracy = 0.01)),
                            Portfolio_3 = c(scales::percent(c(means_sp$p3, std_sp$p3), accuracy = 0.01)),
                            Portfolio_4 = c(scales::percent(c(means_sp$p4, std_sp$p4), accuracy = 0.01)),
                            Portfolio_5 = c(scales::percent(c(means_sp$p5, std_sp$p5), accuracy = 0.01)))

rownames(summary_sp_ch) <- c("Mean", "Std")


#forward discount calculation

forward_discount_1 <- c()
forward_discount_2 <- c()
forward_discount_3 <- c()
forward_discount_4 <- c()
forward_discount_5 <- c()

for(i in 1:(nrow(forward_discount))){
  ranks[i,] <- rank(as.numeric(forward_discount[i,]), ties.method = "average")
  forward_discount_1[i] <- mean(as.numeric(forward_discount[i, as.vector(ranks[i,]) < 3]))
  forward_discount_2[i] <- mean(as.numeric(forward_discount[i, as.vector(ranks[i,]) < 5 & 2 < as.vector(ranks[i,])]))
  forward_discount_3[i] <- mean(as.numeric(forward_discount[i, as.vector(ranks[i,]) < 7 & 4 < as.vector(ranks[i,])]))
  forward_discount_4[i] <- mean(as.numeric(forward_discount[i, as.vector(ranks[i,]) < 9 & 6 < as.vector(ranks[i,])]))
  forward_discount_5[i] <- mean(as.numeric(forward_discount[i, as.vector(ranks[i,]) > 8]))
}

forward_discount1 <- xts(forward_discount_1, order.by=index(forward_discount))[-nrow(forward_discount)]
forward_discount2 <- xts(forward_discount_2, order.by=index(forward_discount))[-nrow(forward_discount)]
forward_discount3 <- xts(forward_discount_3, order.by=index(forward_discount))[-nrow(forward_discount)]
forward_discount4 <- xts(forward_discount_4, order.by=index(forward_discount))[-nrow(forward_discount)]
forward_discount5 <- xts(forward_discount_5, order.by=index(forward_discount))[-nrow(forward_discount)]



means_fo <- data.frame(p1 = c(mean(forward_discount1, na.rm=TRUE) * 12),
                       p2 = c(mean(forward_discount2, na.rm=TRUE) * 12),
                       p3 = c(mean(forward_discount3, na.rm=TRUE) * 12),
                       p4 = c(mean(forward_discount4, na.rm=TRUE) * 12),
                       p5 = c(mean(forward_discount5, na.rm=TRUE) * 12))

std_fo <- data.frame(p1 = c(sd(forward_discount1, na.rm=TRUE) * sqrt(12)),
                     p2 = c(sd(forward_discount2, na.rm=TRUE) * sqrt(12)),
                     p3 = c(sd(forward_discount3, na.rm=TRUE) * sqrt(12)),
                     p4 = c(sd(forward_discount4, na.rm=TRUE) * sqrt(12)),
                     p5 = c(sd(forward_discount5, na.rm=TRUE) * sqrt(12))) 

summary_fo_di <- data.frame(Portfolio_1 = c(scales::percent(c(means_fo$p1, std_fo$p1), accuracy = 0.01)),
                            Portfolio_2 = c(scales::percent(c(means_fo$p2, std_fo$p2), accuracy = 0.01)),
                            Portfolio_3 = c(scales::percent(c(means_fo$p3, std_fo$p3), accuracy = 0.01)),
                            Portfolio_4 = c(scales::percent(c(means_fo$p4, std_fo$p4), accuracy = 0.01)),
                            Portfolio_5 = c(scales::percent(c(means_fo$p5, std_fo$p5), accuracy = 0.01)))

rownames(summary_fo_di) <- c("Mean", "Std")

#for pca

port_pca <- merge(portfolio1, portfolio2, portfolio3, portfolio4, portfolio5)

pca <- prcomp(port_pca[-1,])
summary(pca)
