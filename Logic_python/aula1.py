#pt1
"""
DocString
E escrever o que eu
quiser
asdfasdfd
"""

''' Usar para escrever suas notas '''

# Permite escrever um comentário
print(123)  # Na frente
# Abaixo
print(456)



#pt2
# \r\n -> CRLF
# \n -> LF
print(12, 34, 1011, sep="", end='#')
print(56, 78, sep='-', end='\n')
print(9, 10, sep='-', end='\n')



#pt3
"""
DocString
Python = Linguagem de programação
Tipo de tipagem = Dinâmica / Forte
str -> string -> texto
Strings são textos que estão dentro de aspas
"""
print(1234)

# Aspas simples
print('Luiz Otávio')
print(1, 'Luiz "Otávio"')

# Aspas duplas
print("Luiz Otávio")
print(2, "Luiz 'Otávio'")

# Escape
print("Luiz \"Otávio\"")

# r
print(r"Luiz \"Otávio\"")



#pt4
# Tipos int e float
# int -> Número inteiro
# O tipo int representa qualquer número
# positivo ou negativo. int sem sinal é considerado
# positivo.
# print(11) # int
# print(-11) # int
# print(0)

# float -> Número com ponto flutuante
# O tipo float representa qualquer número
# positivo ou negativo com ponto flutuante.
# float sem sinal é considerado positivo.
# print(1.1, 10.11)
# print(0.0, -1.5)

# A função type mostra o tipo que o Python
# inferiu ao valor.
print(type('Otávio'))
print(type(0))
print(type(1.1), type(-1.1), type(0.0))



#pt5
# Tipo de dado bool (boolean)
# Ao questionar algo em um programa,
# só existem duas respostas possíveis:
# sim (True) ou não (False).
# Existem vários operadores para "questionar".
# Dentre eles o ==, que é um operador lógico que
# questiona se um valor é igual a outro
print(10 == 10)  # Sim => True (Verdadeiro)
print(10 == 11)  # Não => False (Falso)
print(type(True))
print(type(False))
print(type(10 == 10))
print(type(10 == 11))


