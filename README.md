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

- <span style="color:#dd7538">USERNAME2</span> = Usuário do ***Domínio WEB*** no ***App***

- <span style="color:#dd7538">PASSWORD2</span> = Senha do ***Domínio WEB*** no ***App***

---

Finalizando, o que foi pedido acima, agora é só rodar dentro do *bash* em que foi ativado o *venv*.

```bash
$ python main.py
```
