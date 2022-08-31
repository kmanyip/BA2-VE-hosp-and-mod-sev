#set working directory
rm(list = ls())
library(openxlsx)
library(dplyr)
library(data.table)
library(MASS)
library(tibble)
library(tidyverse)

setwd("D:/vaccination")

df <- read.xlsx("hospitalization and severity dataset.xlsx")

###########################VE - vaccination population ######################

#Sex: 1= M, 2 = F
#Age group: 1: 3-11, 2= 12-18

############Table 1 Characteristics

#calculate total vaccination no
df %>% dplyr::group_by(Sex, Age_gp, Vaccine_type, Dose) %>% 
  summarize(vac.no = max(cum.vac.no),
            pop = max(total.pop),
            hospitalization = sum(hosp.adm, na.rm=T),
            moderate.to.severe = sum(mo.sev, na.rm=T),
            infection = sum(infection.no, na.rm=T)) -> et1

#calculate total vaccination no of the latest does
et1$id <- paste0(et1$Sex, et1$Age_gp, et1$Vaccine_type)
et1[order(et1$id),] -> et1
et1$g <- rleid(et1$id)

table(et1$id)

temp <- NULL
for( i in 1:max(et1$g)){
  tp <- subset(et1, g == i)
  if(NROW(tp)==3){
    tp[1,5] <- tp[1,5]-tp[2,5]
    tp[2,5] <- tp[2,5]-tp[3,5]
  }
  tp -> temp[[i]]
}

et1.1 <- do.call("rbind", temp)

et1.1$id <- paste0(et1.1$Sex, et1.1$Age_gp)
et1.1[order(et1.1$id),] -> et1.1
et1.1$g <- rleid(et1.1$id)

table(et1.1$id)

temp <- NULL
for( i in 1:max(et1.1$g)){
  tp <- subset(et1.1, g == i)
  tp$vac.no <- ifelse(tp$Dose == 0, tp$pop - sum(tp$vac.no, na.rm=T), tp$vac.no)
  tp$vac.per <- round(tp$vac.no/tp$pop,3)
  tp$hosp.per <- round(tp$hospitalization/tp$infection,3)
  tp$mo.sev.per <- round(tp$moderate.to.severe/tp$infection,3)
  tp$infect.per <- round(tp$infection/tp$vac.no,3)
  tp -> temp[[i]]
}

et1.2 <- do.call("rbind", temp)
et1.2 <- dplyr::select(et1.2, -c("id", "g"))

##############Table 2 Vaccine effectiveness
######### Hospitalization
df$log.pop <- log(df$pop.at.risk)
df$vcdose <- with(df, ifelse(Vaccine_type == "BioNTech" & Dose ==1, 1, ifelse(
  Vaccine_type == "Sinovac" & Dose ==1, 2, ifelse(
    Vaccine_type == "BioNTech" & Dose ==2, 3, ifelse( 
      Vaccine_type == "Sinovac" & Dose ==2, 4, ifelse(
        Vaccine_type == "BioNTech" & Dose ==3, 5, ifelse(
          Vaccine_type == "Sinovac" & Dose ==3, 6, 0)))))))

df[order(df$Date),] -> df

df$day.count <- rleid(df$Date)

df$vcdose <- as.factor(df$vcdose)

table(df$vcdose)

df$Vaccine_type <- relevel(as.factor(df$Vaccine_type), ref="None")

#Age group: 3-11
df.g1 <- subset(df, Age_gp == 1 & log.pop != -Inf)
table(df.g1$vcdose, df.g1$hosp.adm)

df.g1 <- subset(df, Age_gp == 1 & log.pop != -Inf & (vcdose ==0|vcdose ==1|vcdose ==2|vcdose ==4))
df.g1$Dose <- as.factor(df.g1$Dose)
m1 <- glm.nb(hosp.adm ~ vcdose + day.count +offset(log.pop), link = log, data=df.g1 )
t.g1 <- as.data.frame(summary(m1)$coefficients)
t.g1 <- tibble::rownames_to_column(t.g1, "coeff")
ci <- as.data.frame(confint(m1, level=0.95))
ci <- tibble::rownames_to_column(ci, "coeff")
t.g1 <- merge(t.g1, ci, by = "coeff")
t.g1$rr <- exp(t.g1$Estimate)
t.g1$rr2.5 <- exp(t.g1$`2.5 %`)
t.g1$rr97.5 <- exp(t.g1$`97.5 %`)
t.g1$ve <- (1- t.g1$rr)
t.g1$ve2.5 <- (1- t.g1$rr2.5)
t.g1$ve97.5 <- (1- t.g1$rr97.5)

t.g1.2 <- dplyr::select(t.g1, c("coeff", "rr"))
t.g1.2$Age_gp = "1"

#Age group: 12-18
df.g2 <- subset(df, Age_gp == 2 & log.pop != -Inf)
table(df.g2$vcdose, df.g2$hosp.adm)

df.g2 <- subset(df, Age_gp == 2 & log.pop != -Inf  & vcdose !=5 & vcdose !=6 )
m2 <- glm.nb(hosp.adm ~ vcdose + day.count +offset(log.pop), link = log, data=df.g2 )
t.g2 <- as.data.frame(summary(m2)$coefficients)
t.g2 <- tibble::rownames_to_column(t.g2, "coeff")
ci2 <- as.data.frame(confint(m2, level=0.95))
ci2 <- tibble::rownames_to_column(ci2, "coeff")
t.g2 <- merge(t.g2, ci2, by = "coeff")
t.g2$rr <- exp(t.g2$Estimate)
t.g2$rr2.5 <- exp(t.g2$`2.5 %`)
t.g2$rr97.5 <- exp(t.g2$`97.5 %`)
t.g2$ve <- (1- t.g2$rr)
t.g2$ve2.5 <- (1- t.g2$rr2.5)
t.g2$ve97.5 <- (1- t.g2$rr97.5)

t.g2.2 <- dplyr::select(t.g2, c("coeff", "rr"))
t.g2.2$Age_gp = "2"

rr.hosp <- rbind(t.g1.2, t.g2.2)
colnames(rr.hosp) <- c("vcdose", "rr", "Age_gp")

##### Moderate-to-severe disease #####
df$log.pop <- log(df$pop.at.risk)
df$vcdose <- with(df, ifelse(Vaccine_type == "BioNTech" & Dose ==1, 1, ifelse(
  Vaccine_type == "Sinovac" & Dose ==1, 2, ifelse(
    Vaccine_type == "BioNTech" & Dose ==2, 3, ifelse( 
      Vaccine_type == "Sinovac" & Dose ==2, 4, ifelse(
        Vaccine_type == "BioNTech" & Dose ==3, 5, ifelse(
          Vaccine_type == "Sinovac" & Dose ==3, 6, 0)))))))

df[order(df$Date),] -> df

df$day.count <- rleid(df$Date)

df$vcdose <- as.factor(df$vcdose)

table(df$vcdose)

df$Vaccine_type <- relevel(as.factor(df$Vaccine_type), ref="None")

#Age group: all
table(df$vcdose, df$mo.sev)
df.g <- subset(df, log.pop != -Inf & (vcdose ==0 | vcdose ==1 |vcdose ==2| vcdose ==3 |vcdose==4))
m1 <- glm.nb(mo.sev ~ vcdose+ day.count +offset(log.pop), link = log, data=df.g )
t.g <- as.data.frame(summary(m1)$coefficients)
t.g <- tibble::rownames_to_column(t.g, "coeff")
ci <- as.data.frame(confint(m1, level=0.95))
ci <- tibble::rownames_to_column(ci, "coeff")
t.g <- merge(t.g, ci, by = "coeff")
t.g$rr <- exp(t.g$Estimate)
t.g$rr2.5 <- exp(t.g$`2.5 %`)
t.g$rr97.5 <- exp(t.g$`97.5 %`)
t.g$ve <- (1- t.g$rr)
t.g$ve2.5 <- (1- t.g$rr2.5)
t.g$ve97.5 <- (1- t.g$rr97.5)

t.g.2 <- dplyr::select(t.g, c("coeff", "rr"))

rr.sev <- t.g.2
colnames(rr.sev) <- c("vcdose", "rr")

#daily expected number
### hospitalization
df %>% dplyr::group_by(Date, vcdose, Age_gp) %>% summarize(tot.hosp.adm= sum(hosp.adm, na.rm=T)) -> ep.case
ep.case$vcdose <- paste0("vcdose", ep.case$vcdose)

ep.case2 <- merge(ep.case, rr.hosp, by = c("vcdose", "Age_gp"), all.x=T)
ep.case2$rr <- ep.case2$rr
ep.case2$rr <- with(ep.case2, ifelse((rr <=0 & !is.na(rr)) | vcdose == "vcdose0", 1, rr))

ep.case3 <- subset(ep.case2, !is.na(rr))

ep.case3 %>% gather(var, value, tot.hosp.adm:rr) -> ep.case4
ep.case4$id <- paste0(ep.case4$vcdose, ep.case4$var)

ep.case4 %>% dplyr::select(-c("var", "vcdose")) %>% spread(id, value) -> ep.case5

ep.case5$vcdose3tot.hosp.adm <- ifelse(is.na(ep.case5$vcdose3tot.hosp.adm), 0, ep.case5$vcdose3tot.hosp.adm )
ep.case5$vcdose3rr <- ifelse(is.na(ep.case5$vcdose3rr), 1, ep.case5$vcdose3rr)

ep.case5$ep.case <- with(ep.case5,vcdose0tot.hosp.adm+vcdose1tot.hosp.adm/vcdose1rr + vcdose2tot.hosp.adm/vcdose2rr + vcdose3tot.hosp.adm/vcdose3rr +vcdose4tot.hosp.adm/vcdose4rr)
ep.case5$tot.case <- with(ep.case5,vcdose0tot.hosp.adm+vcdose1tot.hosp.adm + vcdose2tot.hosp.adm + vcdose3tot.hosp.adm +vcdose4tot.hosp.adm)


### Moderate-to-severe cases
df %>% dplyr::group_by(Date, vcdose, Age_gp) %>% summarize(tot.mo.sev= sum(mo.sev, na.rm=T)) -> ep.case
ep.case$vcdose <- paste0("vcdose", ep.case$vcdose)

ep.case2 <- merge(ep.case, rr.sev, by = c("vcdose"), all.x=T)
ep.case2$rr <- ep.case2$rr
ep.case2$rr <- with(ep.case2, ifelse((rr <=0 & !is.na(rr)) | vcdose == "vcdose0", 1, rr))

ep.case3 <- subset(ep.case2, !is.na(rr))

ep.case3 %>% gather(var, value, tot.mo.sev:rr) -> ep.case4
ep.case4$id <- paste0(ep.case4$vcdose, ep.case4$var)

ep.case4 %>% dplyr::select(-c("var", "vcdose")) %>% spread(id, value) -> ep.case5

ep.case5$ep.case <- with(ep.case5,vcdose0tot.mo.sev+vcdose1tot.mo.sev/vcdose1rr + vcdose2tot.mo.sev/vcdose2rr + vcdose3tot.mo.sev/vcdose3rr +vcdose4tot.mo.sev/vcdose4rr)
ep.case5$tot.case <- with(ep.case5,vcdose0tot.mo.sev+vcdose1tot.mo.sev + vcdose2tot.mo.sev + vcdose3tot.mo.sev +vcdose4tot.mo.sev)
