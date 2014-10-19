### inscriPET  
 
#### Sistema de cadastro de alunos para o projeto [sisPET](https://github.com/wsilverio/sisPET)
  
 **Licença:** [Atribuição CC BY](http://creativecommons.org/licenses/by/4.0/)  
  
      
**Dependências:**

 * [Google Spreadsheets Python API](https://github.com/burnash/gspread)
 * [ZBar bar code reader](http://zbar.sourceforge.net/)
 * Webcam

#### Como usar:
 * Criar o banco de dados em [g.co/oldsheets](https://g.co/oldsheets)
 * Copiar a chave (key) da planilha (URL)
 * Substituir os dados do arquivo [PET.py](https://github.com/wsilverio/inscriPET/blob/master/PET.py)
 * Verificar o caminho da webcam:  
  
 ```console
$ ls /dev/video*
```  
 * Substituir o endereço da webcam (padrão "/dev/video0") no comando **scanner.init()**,  função **configura_scanner()** no arquivo [inscriPET.py](https://github.com/wsilverio/inscriPET/blob/master/inscriPET.py)
 * Executar o programa:  
  
 ```console
$ python inscriPET.py  
```
