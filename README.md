# Email_automatico_Folhas_Ponto
 Código que envia folhas de ponto para funcionários de forma automática.

Trabalho com um sistema de geração de folhas de ponto, só que o sistema gera um arquivo único, em PDF, com todas as folhas de ponto juntas. O cliente pediu para que fosse criado um arquivo de folha de ponto para cada funcionário e que este arquivo fosse enviado, como anexo, cada um para o funcionário proprietário da folha de ponto.

Para atender a solicitação, utilizei Python, PDF, Excel e E-mail.

Na primeira parte do código foram importadas as bibliotecas python utilizadas para o funcionamento, bibliotecas para PDF, Excel, caminhos do computador e E-mail.

Depois da importação das bibliotecas, o código lê e armazena as informações do arquivo de folhas de ponto que o sistema gera.

Após a leitura, o código separa este arquivo em diversas folhas de ponto, faz uma análise para obter as informações do nome de cada funcionário e salva os novos arquivos em uma pasta no computador, com o nome de cada pessoa.

Com as folhas de ponto salvas no pc, o código faz uma nova análise, agora ele compara cada nome de arquivo dentro de uma planilha em Excel. Nesta planilha é preciso que existam duas colunas, com o nome do funcionário e o e-mail dele. Finalizando a análise, é armazenada a informação de qual folha de ponto pertence a qual e-mail.

Para finalizar, o código cria um e-mail com cada folha de ponto e faz o envio. Este e-mail é o da pessoa responsável para enviar os dados.

Cada funcionário irá receber em seu e-mail a sua respectiva folha de ponto.

