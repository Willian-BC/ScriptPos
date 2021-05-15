library(readxl)
library(tidyr)
library(dplyr)
#abrindo os dados
FIIs <- read_excel("C:/Users/willi/Desktop/FIIs.xls")
#criando um data frame clone para alteração
dados <- as.data.frame(FIIs)
#Separar dados por colunas no data frame
dados <- dados %>% separate(Cotacao, c("R$","valor"), sep=3)
dados <- dados[,-2] #excluir coluna do data frame

#alterando os valores vazios para zero dentro da coluna separada
for (i in 1:length(dados[,3])){
  if ((dados[i,3]) == "") {
    dados[i,3] <- 0
  }
}

str(dados[7,3])
dados$valor <- as.numeric(dados$valor)
