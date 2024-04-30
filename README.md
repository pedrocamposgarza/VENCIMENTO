# VENCIMENTO

### Garza Inteligência Financeira
## Projeto Vencimento


### *Pedro Henrique Campos Moreira*
---

## 1. Instale o Visual Studio Code (Vscode)

  - **Entre na Microsoft Store**  
  - **Pesquise Visual studio Code**
  - **Instale a IDE**

## 2. Instale o Python (Versão mais recente e compatível com o seu S.O.)

   ```shell
   https://www.python.org/downloads/
   ```

## 3. Clone este repositório:

  - **Copie o arquivo projeto27.py**  
  - **Abra o Vscode**
  - **Crie um novo arquivo.py (file)**

## 4. Bibliotecas - Importações

**Bibliotecas devem ser baixadas no terminal.**

   ```shell
  pip install openpyxl
   ```

   ```shell
  conda install -c anaconda tk
   ```

   ```shell
   pip install pypiwin32
   ```

   ```shell
  pip install icalendar

   ```

## 5. Preparo

  - **Cole o código no novo arquivo criado**  
  - **Se necessário, adicione o arquivo clientes.xlxs em sua pasta**
  - **Ter os arquivos pdf's e excel baixados em sua máquina**

## 6. Execução

 1. **Compile o código**  
 2. **Selecione o arquivo excel do cliente**
 3. **Verifique seu email**
 4. **Clique na seta direita do inside **
 5. **Adicione ao seu calendario**
   
## 7. Observações

```
    # Enviar o convite de calendário por e-mail
    mail = outlook.CreateItem(0)  # 0 para e-mail
    mail.Subject = "Lembrete de Eventos"
    mail.Body = "Por favor, adicione este convite ao seu calendário."
    # Insira o endereço de e-mail do destinatário aqui
    mail.To = "pedro.moreira@garzaif.com.br"
    mail.Attachments.Add(temp_file.name, 1)  # Anexar o convite de calendário
```
 - Modificar o email acima para o seu , para que receba o inside necessario  


 ## 8. Recomandações
  - Em caso de erro , fechar todas as jenals e iniciar novamente o codigo
  - Todos os arquivos .xlxs devem estar no mesmo padrão, Coluna das datas de vencimento sendo  Q e colunas dos nomes sendo A
  
---
