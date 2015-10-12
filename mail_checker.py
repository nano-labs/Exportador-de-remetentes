#!/usr/bin/env python
# -*- encoding: utf-8 -*-
u"""
Exportador de remetentes.

Rabisco 100% GoHorse para exportar uma lista com todos
os remetentes de emails de uma conta e gerar um arquivo
de import do gmail.
"""


import sys
from datetime import datetime
import imaplib
import email
from email.mime.text import MIMEText
import csv

print "\nDigite o endereço de email:"
username = raw_input()
print "\nDigite a senha:"
password = raw_input()

try:
    s = imaplib.IMAP4_SSL('imap.gmail.com')
    s.login(username, password)
except:
    print u"Senha errada"
    sys.exit()
s.select("INBOX")

status, ids = s.search(None, "All")
ids = ids[0].split(" ")
total = len(ids)
atual = 0
toda_lista = []
print "Iniciando busca de contatos:"
for i in ids[:100]:
    atual += 1
    # import ipdb;ipdb.set_trace()
    mensagem = email.message_from_string(s.fetch(i, '(RFC822)')[1][0][1])
    # assunto = mensagem.get("Subject")
    remetente = mensagem.get("From")
    if "<" in remetente:
        try:
            endereco = remetente.split("<")[1].split(">")[0]
            nome = remetente.split("<")[0]
        except:
            print "Nao consegui entender o email '%s'" % remetente
    else:
        endereco = remetente
        nome = remetente
    toda_lista.append((nome, endereco))
    print u"%s de %s: %s" % (atual, total, remetente)

toda_lista = list(set(toda_lista))

saida = open("lista_gmail.csv", "w")
cabecalho = "Name,Given Name,Additional Name,Family Name,Yomi Name,Given Name Yomi,Additional Name Yomi,Family Name Yomi,Name Prefix,Name Suffix,Initials,Nickname,Short Name,Maiden Name,Birthday,Gender,Location,Billing Information,Directory Server,Mileage,Occupation,Hobby,Sensitivity,Priority,Subject,Notes,Group Membership,E-mail 1 - Type,E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value\n"
linha = "%s,,,,,,,,,,,,,,,,,,,,,,,,,,,* ,%s,,\n"
saida.write(cabecalho)
for i in toda_lista:
    saida.write(linha % i)
saida.close()
