# Supervisório Estação de Furos (Grupo 5): plano e correções

Base: rubrica ECAA02 2026, documentação da estação FESTO (UNIFEI) e análise do `Grupo5_Furo.prj` / `Estação_furos.lib` em 23/11/2026.

## 1. Situação atual por critério da rubrica

| Critério (peso) | Já existe | Falta | Estimativa hoje | Meta |
|---|---|---|---|---|
| 1. Telas e variáveis (20%) | Sinótico com mesa, furadeira, pistão; 9 de 20 tags | Botões XControl sem escrita, QUIT, 7 sensores, 3 saídas, intertravamentos, unidades | 3–4 | 10 |
| 2. Telas/quadros, usuários, galeria (10%) | TelaInicial, TelaMenu, TelaPrincipal, Quadro1; biblioteca de XControls | Usuários e permissões; quadro com navegação completa | 1–2 de 3 | 3 de 3 |
| 3. Alarmes, históricos, consultas, gráficos, relatórios (20%) | Só a TelaAlarmes vazia | Os 5 recursos | 0 de 5 | 5 de 5 |
| 4. Scripts, XControls, XObjects, fórmulas (10%) | Scripts de navegação e comandos; XControls com bug | XControls corrigidos, XObject, Fórmula (receita) | 1–2 de 4 | 4 de 4 |
| 5. Demonstração (20%) | Driver SNP configurado | Cenário de alarme/falha ensaiado | – | 10 |
| 6. Arguição (20%) | – | Divisão de responsabilidades e roteiro | – | 10 |

Limite de 20 tags: 6 comandos + 7 sensores + 5 saídas + 2 contadores = 20. Hoje existem 9.

## 2. Correções imediatas (botões XControl)

Diagnóstico:
- No XControl `Botao`, o script `Circulo3_MouseDown` faz `Application.SetValue "Apertado", CLng(1)`. "Apertado" não é tag, então nada é escrito, e nada volta a 0.
- `XControl11.Apertado` está ligado a `:Driver1.start.Value` com link simples (só leitura).
- `XControl12` (Reset), `XControl13` (Position) e `XControl15` (Stop) não têm link.
- `ToggleButton1` (Auto) tem script `Application.SetValue "", CLng(0)` (tag vazio) e link com Inverter marcado.

Correção:
1. Biblioteca `Estação_furos.lib` → XControl `Botao` → objeto `Circulo3`:
   ```vb
   Sub Circulo3_MouseDown(Button, ShiftState, MouseX, MouseY)
       Botao.Apertado = True
   End Sub

   Sub Circulo3_MouseUp(Button, ShiftState, MouseX, MouseY)
       Botao.Apertado = False
   End Sub
   ```
   Salvar a lib e atualizar as instâncias.
2. Na TelaPrincipal, link **Bidirecional** de `Apertado` em cada instância: Start → `Driver1.start.Value`, Stop → `Driver1.stop.Value`, Position → `Driver1.position.Value`, Reset → `Driver1.reset.Value`, Quit → `Driver1.quit.Value` (criar).
3. XControl `Chave`: `Ligado` com link Bidirecional em `Driver1.auto_manual.Value`; no clique: `Chave.Ligado = Not Chave.Ligado`.
4. Apagar o script do ToggleButton1 e desmarcar Inverter (ou apagar o botão e usar só a Chave).
5. Testar com o Tag Monitor: o tag deve ir a 1 e voltar a 0.

## 3. Tags: 20 variáveis (limite do driver)

Códigos N2 do manual do driver GE SNP (SNP_BR.pdf), todos em bit: 0 = %I, 2 = %Q, 6 = %M; 20 = %R (word). N1 = 0 (CLP 12345 no GEFANUC.ID), N4 = 0.

Driver: P1 = 1 (COM1), P2 = 7 (19200 bps, CPU 90-30), P3 = 50 ms, P4 = 20 (0,2 s). No laboratório, conferir em qual COM o cabo aparece.

| # | Tag | N2 | N3 | Endereço | Significado |
|---|---|---|---|---|---|
| 1 | start | 6 | 300 | %M300 | Botão Start |
| 2 | stop | 6 | 301 | %M301 | Botão Stop |
| 3 | position | 6 | 302 | %M302 | Botão Position |
| 4 | quit (**criar**) | 6 | 303 | %M303 | Botão Quit |
| 5 | reset | 6 | 304 | %M304 | Botão Reset |
| 6 | auto_manual | 6 | 305 | %M305 | Chave Auto/Manual |
| 7 | emergencia (**criar**) | 0 | 8 | %I8 | Emergência (NF: 0 = acionada) |
| 8 | sens_mesa (**criar**) | 0 | 9 | %I9 | Sensor indutivo da mesa |
| 9 | fur_cima (**criar**) | 0 | 10 | %I10 | Furadeira em cima |
| 10 | fur_baixo (**criar**) | 0 | 11 | %I11 | Furadeira embaixo |
| 11 | teste_baixo (**criar**) | 0 | 13 | %I13 | Teste embaixo (peça furada) |
| 12 | peca_travada (**criar**) | 0 | 15 | %I15 | Trava avançada |
| 13 | presenca (**criar**) | 0 | 16 | %I16 | Peça na entrada |
| 14 | Furadeira | 2 | 1 | %Q1 | Motor da furadeira |
| 15 | Motor_mesa | 2 | 2 | %Q2 | Motor da mesa |
| 16 | sol_desce_fur (**criar**) | 2 | 3 | %Q3 | Solenoide descida furadeira |
| 17 | sol_teste (**criar**) | 2 | 5 | %Q5 | Solenoide descida do teste |
| 18 | sol_trava (**criar**) | 2 | 6 | %Q6 | Solenoide da trava |
| 19 | CtPcFurado | 20 | 100 | %R100 | Contador peças furadas |
| 20 | CtPcNFurado | 20 | 104 | %R104 | Contador peças não furadas |

Fora do plano (repetem informação de outro tag):
- %Q4 sobe furadeira: subida = Furadeira ligada, sol_desce_fur desligado e fur_cima desligado.
- %I12 teste em cima: haste em cima = teste_baixo e sol_teste desligados.
- %I14 peça liberada: liberada = peca_travada e sol_trava desligados.

Valores calculados (TotalPecas, TaxaRefugo, TempoCiclo, Estado, ComFalha, FalhaFuradeira, MetaProducao, LimiteRefugo, TempoMaxCiclo) ficam como propriedades do XObject ControleProducao, não como tags.

## 4. Plano por recurso

### Critério 1: telas e variáveis
- Ligar cada sensor e cada saída a uma animação no sinótico: furadeira desce com %I11 / %Q3, trava com %Q6 / %I15, haste de teste com %Q5 / %I13, peça na entrada com %I16, mesa gira com %Q2.
- Painel de sinais (LEDs) com os 9 sensores e as 6 saídas.
- Intertravamentos na IHM: Position habilitado só em Manual; Start e Quit só em Auto; Reset só para Supervisor e com confirmação na própria tela; Stop sempre habilitado. Usar a propriedade `Enabled` com link.
- Unidades em todos os displays: "peças", "%", "s".

### Critério 2: estrutura do E3
- **Telas e quadros**: Quadro com cabeçalho (estado, comunicação, usuário, hora, último alarme), área central e menu inferior (Sinótico, Alarmes, Histórico, Relatórios, Receitas, Usuários, Sair).
- **Usuários e permissões**: grupos Operador, Supervisor e Manutenção (ver matriz abaixo). Login na TelaInicial com `Application.Login(True)`. Checar grupo com `Application.IsUserMemberOfGroup("Supervisor")` antes do Reset e das receitas.
- **Galeria**: XControls na `Estação_furos.lib`, mais símbolos da Galeria do E3 (motor, sensor, cilindro).

| Ação | Operador | Manutenção | Supervisor |
|---|---|---|---|
| Start / Stop / Quit | sim | sim | sim |
| Position e modo Manual | não | sim | sim |
| Reset dos contadores | não | não | sim |
| Reconhecer alarmes | sim | sim | sim |
| Gerar relatório | sim | sim | sim |
| Carregar ou salvar receita | não | não | sim |

### Critério 3: alarmes, históricos, consultas, gráficos, relatórios
1. **Alarmes**: Servidor de Alarmes + Configuração de Alarmes, com áreas `Processo` e `Sistema`:
   - Digital `emergencia = 0` → "Emergência acionada", severidade Alta.
   - Digital `ControleProducao.ComFalha = 1` → "Falha de comunicação com o CLP", Alta. `ComFalha` ligado a `Driver1.Motor_mesa.Quality <> 192`.
   - Digital `ControleProducao.FalhaFuradeira = 1` → "Furadeira não atingiu fim de curso em 5 s", Alta.
   - Analógico `ControleProducao.TaxaRefugo`: Hi 10 % (Média), HiHi 20 % (Alta). **Limite com prioridade.**
   - Analógico `ControleProducao.TempoCiclo`: Hi 20 s (Baixa), HiHi 30 s (Média).
   - Tela de alarmes com E3Alarm e botões `AckSelected` / `AckCurrentFilter`. Alarmes gravados no banco.
2. **Histórico**: DBServer (Access `.mdb`) + objeto Histórico com CtPcFurado, CtPcNFurado, TotalPecas, TaxaRefugo, TempoCiclo e Estado. Gravar a cada 10 s ou por banda morta.
3. **Consulta**: E3Browser com Consulta filtrada por data inicial e final (mesmo padrão do `spDataInicial` / `spDataFinal` do treinamento).
4. **Gráfico**: E3Chart histórico com peças furadas e não furadas ao longo do turno, mais uma pena de taxa de refugo no eixo secundário.
5. **Relatório**: E3Reports com consulta ao histórico: produção por hora, totais, taxa de refugo e lista de alarmes do período.
   ```vb
   Set rel = Application.LoadReport("[RelProducao]")
   rel.Query(0).SetVariableValue "DataIni", Screen.Item("spDataInicial").Value
   rel.Query(0).SetVariableValue "DataFim", Screen.Item("spDataFinal").Value
   rel.PrintPreview
   ```

### Critério 4: scripts, XControls, XObjects, fórmulas
- **Scripts**: comandos, login, relatório e timeout da furadeira.
- **XControls**: Botao, Chave, Mesa, Furadeira e Pistao, todos ligados a tags reais.
- **XObject `ControleProducao`** (no servidor): recebe furadas, não furadas e as saídas; calcula TotalPecas, TaxaRefugo, TempoCiclo e Estado ("Manual", "Aguardando peça", "Girando mesa", "Furando", "Testando", "Parado"); dispara FalhaFuradeira se %Q3 ficar ligado mais de 5 s sem %I11.
- **Fórmula (receita) `Parametros`**: unidades MetaProducao, LimiteRefugo e TempoMaxCiclo; conjuntos "Padrão" (60 / 10 / 20), "Lote de teste" (10 / 30 / 40) e "Alta exigência" (80 / 5 / 15). Os limites dos alarmes vêm desses tags.

### Critério 5: roteiro de demonstração (cerca de 8 min)
1. Login como Operador; mostrar o quadro e a comunicação OK.
2. Modo Manual (Manutenção): Position gira 90°.
3. Modo Auto: Start, colocar peça, ciclo completo animado, contador sobe.
4. **Cenário de alarme**: apertar a emergência → alarme Alto no banner → reconhecer → normalizar. (Alternativa: peça sem furo → refugo passa de 10 % → alarme Médio.)
5. Histórico e gráfico do período; gerar relatório.
6. Supervisor: Reset dos contadores e carga de receita.

### Critério 6: arguição
- Dividir por área: (A) driver, tags, sinótico e XControls; (B) alarmes, histórico, consultas e gráficos; (C) relatórios, usuários, XObject e fórmulas. Cada um deve saber explicar o sistema inteiro.
- Pontos de metodologia de IHM (ISA-101): fundo cinza neutro, cor só para estado anormal, prioridade de alarme com cor e forma, hierarquia de telas (visão geral → sinótico → detalhe), consistência de posição dos botões, intertravamento e confirmação de ações críticas.
- Melhorias que podem propor: OEE, alarme por tempo de ciclo médio móvel, acesso via E3 WebViewer, receita por lote.
