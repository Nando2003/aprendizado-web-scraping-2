# *RPA_2*

Este *BOT* feito para a extração de arquivos do *Dominio WEB* e importa-los para **EFD Contribuições**.

# Como instalar ?

Primeiramente, clone esse repositorio em um diretorio da sua escolha no seu computador.

```bash
$ git clone <url_do_repositorio.git>
```

Agora, crie um ambiente virtual (*venv*) dentro do diretorio da **RPA_2**.

```bash
$ python -m venv venv
```

A partir da criação de uma pasta *venv*, você deve ativar o ambiente virtual para que possa instalar as dependências e rodar o *BOT*.

```bash
$ source venv/Scripts/activate
```

Após a ativação do *venv*, você perceberá que todos os comandos que rodados dentro desse *bash* terá no final um ```(venv)```. Isso quer dizer que você ativou o *venv* com sucesso.

Segundamente, instale as dependências que estarão dentro do *requirements.txt*:

```bash
$ pip install -r requirements.txt
```

# Como rodar o *BOT* ?

Antes de começarmos os trabalhos, devemos fazer algo muito importante: criar um arquivo *.env* que deverá seguir o exemplo que está em *.env_example*.

---

- <span style="color:#dd7538">USERNAME1</span> = Usuário do ***Domínio WEB*** na ***WEB***

- <span style="color:#dd7538">PASSWORD1</span> = Senha do ***Domínio WEB*** na ***WEB***

---

- <span style="color:#dd7538">USERNAME1</span> = Usuário do ***Domínio WEB*** no ***App***

- <span style="color:#dd7538">PASSWORD1</span> = Senha do ***Domínio WEB*** no ***App***

---

Finalizando, o que foi pedido acima, agora é só rodar dentro do *bash* em que foi ativado o *venv*.

```bash
$ python main.py
```

# Tecnologias Usadas:

Nesse projeto de *web scraping*, as tecnologias usadas para que todo o processo fosse realizado foram:

- ***Selenium***: Usado bastante para testes automatizados, nesse projeto foi usado como um facilitador para que o ***Domínio WEB*** pudesse ser aberto.

- ***Pyautogui***: Biblioteca famosa no ramo das automações em *Python*. Ela foi usada especialmente para identificar imagens fora do navegador e automatizar processos a partir delas.

- ***Openpyxl***: Especialmente, feita para mexer em arquivos <span style="color:#800080">*.xlsx*</span>. Ela foi indispensavel para disponibilizar para o usuário uma resposta visual de onde os arquivos foram baixados e os status das empresas no sistema.

- ***Dotenv***: Por último e não menos especial, a biblioteca que permite a criação de variáveis de ambiente (geralmente, são informações importantes e confidencias). Ela permite que o arquivo *.env* seja carregado e possamos usar as variáveis em nosso código. Como podemos ver no <span style="color:#800080">*main.py*</span> :
```python
# Imports...

if __name__ == "__main__":
    """
    Pega o diretorio que este arquivo se encontra + dotenv_files/.env
    """
    dotenv_path = os.path.join(os.path.dirname(__file__), "dotenv_files/.env") 
    load_dotenv(dotenv_path)
    
    # Código...
    
    dominio_web = DominioWeb(
        download_path=companies_path,
        excel_path=excel_path,
        username_web=os.environ.get("USERNAME1"),
        password_web=os.environ.get("PASSWORD1"),
        username_local=os.environ.get("USERNAME2"),
        password_local=os.environ.get("PASSWORD2")
    )

    # Código...
```

---