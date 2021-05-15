library(readxl)
library(tidyr)
library(dplyr)
#abrindo os dados
FIIs <- read_excel("C:/Users/willi/Desktop/FIIs.xls")
#criando um data frame clone para alteração
dados <- as.data.frame(FIIs)

#Separar dados por colunas no data frame
#dados <- dados %>% separate(Cotacao, c("R$","valor"), sep=3)
#dados <- dados[,-2] #excluir coluna do data frame

#alterando os valores vazios para zero dentro da coluna separada
#for (i in 1:length(dados[,2])){
#  if ((dados[i,2]) == "") {
#    dados[i,2] <- 0
#  }
#}

#substituindo os dados texto para numero
#str(dados[1,2])
#dados$valor[] <- sub("\\.","",dados$valor[])
#dados$valor <- sub(",",".",dados$valor)
#dados$valor <- as.numeric(dados$valor)
#str(dados[1,2])

#Limpeza dos caracteres 
for (i in c(2:21)) {
  dados[,i] <- sub("\\.","",dados[,i])
  dados[,i] <- sub("\\.","",dados[,i])
  dados[,i] <- sub("\\.","",dados[,i])
  dados[,i] <- sub("\\,",".",dados[,i])
  dados[,i] <- sub("\\+","",dados[,i])
  dados[,i] <- sub("\\R","",dados[,i])
  dados[,i] <- sub("\\$","",dados[,i])
  dados[,i] <- sub("\\N/A",0,dados[,i])
  dados[,i] <- sub("\\%","",dados[,i])
  #dados[,i] <- sub("\\ mi","",dados[,i])
  #dados[,i] <- sub("\\ bi","",dados[,i])
  if (i < 7){
    dados[,i] <- as.numeric(dados[,i])
  }else if (i > 7 & i < 17){
    dados[,i] <- as.numeric(dados[,i])
  }else if (i > 18){
    dados[,i] <- as.numeric(dados[,i])
  }
}

#Tratamento percentual
dados[,3] <- dados[,3]/100
dados[,6] <- dados[,6]/100
dados[,9] <- dados[,9]/100
dados[,16] <- dados[,16]/100

