O objeto AutoTester, pertencente à biblioteca Manuttester.lib, é utilizado para automação de testes e geração de logs.

Propriedades

GerarLogErrosScript

* Descrição: Controla a geração de logs de erro em arquivos .txt.
* Usabilidade:
    * Quando definido como True: O objeto irá gerar arquivos de log (.txt) que contêm informações detalhadas sobre erros ocorridos durante a execução dos scripts.
    * Quando definido como False: Apenas os logs dos testes são registrados, sem incluir os erros de script.
    * Recomendação de uso: Ative esta propriedade durante sessões de depuração ou quando precisar verificar se algum objeto não está sendo analisado.

DebugMode

* Descrição: Controla a geração de arquivos de saída.
* Usabilidade:
    * Quando definido como True: Impede a criação de arquivos de logs (Excel ou .txt).
    * Quando definido como False: Permite a geração de arquivos de logs.
    * Recomendação de uso: Ative esta propriedade durante a fase de desenvolvimento ou testes preliminares quando não desejar gerar arquivos de logs.

PastaParaSalvarLogs

* Descrição: Define o diretório onde os arquivos de log serão salvos.
* Usabilidade:
    * Se preenchido: Os logs serão salvos no caminho especificado.
    * Se deixado em branco: Os logs serão salvos no diretório atual da biblioteca.
    * Formato esperado: Caminho completo (ex: "C:\Logs\MeuProjeto").
    * Recomendação de uso: Defina um diretório específico para facilitar a organização e posterior análise dos logs.

PathNameTelas

* Descrição: Especifica quais telas devem ser incluídas na análise.
* Usabilidade:
    * Lista os PathNames das telas que serão analisadas durante a execução do teste. Para verificar mais de uma tela, separe-os com “/”. 
    * Recomendação de uso: Caso queira verificar todas as telas da aplicação, deixe esta propriedade vazia.
