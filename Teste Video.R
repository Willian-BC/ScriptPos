#importar biblioteca
library(tidyr)
library(dplyr)
#abrir os dados
Fonte_dados <- read.table("C:/Users/willi/Documents/R/teste.txt", 
                          header=FALSE, sep=";")
#transpor os dados para linha
dados <- as.data.frame(t(Fonte_dados))

#separar os dados geral
dados <- dados %>% separate(V1,c("data","preco","qtde"),",")

#separar os a hora
dados <- dados %>% separate(data, c("data","hora"), sep=12)

#separar os dados de data
dados <- dados %>% separate(data, c("mes","dia","ano")," ")

#modificacao do texto para numero
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

#uniao dos dados numericos formando uma data
dados <- dados %>% unite(data,c("dia","mes","ano"),sep="-")

#quebra do dataframe em dois
dados1 <- select(dados[1:2078,],-hora)
dados2 <- dados[2079:2802,]

#separacao e exclusao dos dados incorretos na hora
dados2 <- dados2 %>% separate(hora,c("hora","resto"), " ") %>% select(-resto)
dados2$x <- c("00")
dados2 <- unite(dados2,hora, c("hora","x"), sep="")

#transformacao dos dados para numericos
dados2$preco <- as.numeric(dados2$preco)
dados2$qtde <- as.numeric(dados2$qtde)

#calculo do preco medio com o comando mutate
dados2 <- dados2 %>%  mutate(preco_med = preco / qtde)
