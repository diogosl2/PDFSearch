Busca de Termos em Arquivos


Este projeto Python permite buscar um termo específico dentro de arquivos nos formatos PDF, DOCX, DOC e TXT em um diretório selecionado, contando a quantidade de ocorrências e salvando os resultados em um arquivo Excel (.xlsx).
Funcionalidades
•	Busca recursiva em um diretório selecionado por arquivos nos formatos PDF, DOCX, DOC e TXT.
•	Contagem do número de vezes que o termo buscado aparece em cada arquivo.
•	Captura das linhas completas onde o termo é encontrado nos arquivos DOCX, DOC e TXT.
•	Criação de um arquivo Excel com os resultados contendo:
o	Nome do arquivo.
o	Tipo do arquivo (PDF, DOCX/DOC ou TXT).
o	Termo buscado.
o	Quantidade de ocorrências do termo.
o	Linhas completas onde o termo foi encontrado.
Requisitos
•	Python 3.x
•	Bibliotecas Python:
o	openpyxl
o	fitz (PyMuPDF)
o	docx
o	tqdm
Como Usar
1.	Clone o repositório:
bash
Copiar código
git clone https://github.com/seu_usuario/nome-do-repositorio.git
cd nome-do-repositorio
2.	Instale as dependências:
bash
Copiar código
pip install -r requirements.txt
3.	Execute o script:
o	Execute o script Python file_search.py e siga as instruções para selecionar o diretório e digitar o termo a ser buscado nos arquivos.
bash
Copiar código
python file_search.py
4.	Resultados:
o	Após a execução, um arquivo Excel será gerado na pasta do projeto contendo os resultados da busca.
Exemplo de Estrutura de Arquivos
perl
Copiar código

nome-do-repositorio/
│
├── file_search.py          # Script principal para busca de termos em arquivos
├── README.md               # Documentação do projeto (este arquivo)
└── requirements.txt        # Lista de bibliotecas Python necessárias

Contribuição

Contribuições são bem-vindas! Se você encontrar problemas ou tiver sugestões de melhorias, sinta-se à vontade para abrir uma issue ou enviar um pull request.
Licença

Este projeto está licenciado sob a MIT License.

