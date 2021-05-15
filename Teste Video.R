#importar biblioteca
library(tidyr)
library(dplyr)
#abrir os dados
Fonte_dados <- read.table("C:/Users/willi/Desktop/teste.txt", 
                          header=FALSE, sep=";")
#transpor os dados para linha
dados <- as.data.frame(t(Fonte_dados))
#separar os dados geral
dados <- dados %>% separate(V1,c("data","preço","qtde"),",")
#separar os a hora
dados <- dados %>% separate(data, c("data","hora"), sep=12)
#separar os dados de data
dados <- dados %>% separate(data, c("mês","dia","ano")," ")

for (x in 1:2802) {
  if(dados[x,1]=="Jan") dados[x,1] <- "01"
  if(dados[x,1]=="Feb") dados[x,1] <- "02"
  if(dados[x,1]=="Mar") dados[x,1] <- "03"
  if(dados[x,1]=="Apr") dados[x,1] <- "04"
  if(dados[x,1]=="May") dados[x,1] <- "05"
  if(dados[x,1]=="Jun") dados[x,1] <- "06"
  if(dados[x,1]=="Jul") dados[x,1] <- "07"
  if(dados[x,1]=="Aug") dados[x,1] <- "08"
  if(dados[x,1]=="Sep") dados[x,1] <- "09"
  if(dados[x,1]=="Oct") dados[x,1] <- "10"
  if(dados[x,1]=="Nov") dados[x,1] <- "11"
  if(dados[x,1]=="Dec") dados[x,1] <- "12"
}

dados <- dados %>% unite(data,c("dia","mês","ano"),sep="-")

dados1 <- select(dados[1:2078,],-hora)
dados2 <- dados[2079:2802,]

dados2 <- dados2 %>% separate(hora,c("hora","resto"), " ") %>%
  select(-resto) %>% 
  