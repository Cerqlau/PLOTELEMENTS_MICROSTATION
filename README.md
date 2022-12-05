# PLOTELEMENTS_MICROSTATION_TRACKSURVEY

Este projeto foi desenvolvido em linguagem Visua Basic para automatizar a plotagem de elementos CELL e TEXT no programa Microstation V8i. 
O intuito é facilitar a plotagem dos elementos batimetria, sucata entre outros, em arquivos com extensão DGN. Ele reproduz doi elementos bases selecionados, mantendo o offset entre eles.

## 🚀 Começando

Essas instruções permitirão que você obtenha uma cópia do projeto em operação na sua máquina local para fins de desenvolvimento e teste.

### 📋 Pré-requisitos

```
=> Microstation V8i instalado na máquina. 
=> Arquivo DGN utilizando coodenadas em UTM.
=> Microsoft Excel 2013 ou superior.

```

### 🔧 Pré-configurações

PREPARAR O EXCEL A SER UTILIZADA COMO "BASE DE DADOS": 
1.	Insira na coluna A  todas as coordenadas UTM Leste ;
2.	Insira na coluna B  todas as coordenadas UTM Norte ;
3.	Insira na coluna C  todas as profundidades em acordo com a quantidade de pontos;
4.	Salve e em um local que possa buscar depois, aconselho o desktop para facilitar o acesso. Anexo ao projeto deixei uma planilha modelo “Cotas Batimétricas “. 

Carregamento do projeto: 
1.	Faça o download do arquivo “PLOT_ELEMENTOS_V3” e salve em um local que possa buscar depois, aconselho o desktop para evitar erros com a rede da embarcação;
2.	Com o Microstation carregado vá na aba Utilities > Macro > Project Manager;
3.	Na nova janela que abrira  vá em New Project > copie o arquivo salvo “PLOT_ELEMENTOS_V3” para dentro desta pasta > cancele a ação;
4.	Novamente na janela abra “Load Project” >  selecione o arquivo “PLOT_ELEMENTOS_V3” > Open; 
5.	Na frente do nome do caminho do projeto tem marque a opção “Auto-Load” assim não teremos mais necessidade de executar estes passos novamente.



### ⚙️ Executando o projeto

1.	Previamente plote a primeiro elemento e texto de profundidade, ajuste cor, escala e camada conforme preferir, as propriedades deles servirão de modelos para os demais; 
2.	Utilizando a mesma janela anterior Utilities > Macro > Project Manager, selecione o nome do projeto e click em “Macros”;
3.	Na nova janela selecione “Start.main” > Click em Run para iniciar a aplicação;
4.	Quando o código solicitar carregue a planilha base de dados;
5.	Selecione os elementos a ser copiados e o texto (limitei a somente um objeto cell e um text para facilitar o controle de erros)
6.	Be Happy !!!   


## 📦 Desenvolvimento

Lauro Cerqueira

LinkdIn: https://www.linkedin.com/in/lauro-cerqueira-70473568/

Instagram : @laurorcerqueira

## 📄 Licença

Este projeto está sob a licença MIT - veja o arquivo [LICENSE.md](https://github.com/usuario/projeto/licenca) para detalhes.

## 🎁 Distribuição

* Se gostou conte a outras pessoas sobre este projeto 📢
* Convide alguém da equipe para uma cerveja 🍺 
* Obrigado publicamente 🤓.

