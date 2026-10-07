/* Dados das Diretrizes para a Gestão de Projetos
   Fonte: "Diretrizes para a Gestão de Projetos" (Gerência de Projetos, 121 slides)
   e "Arquitetura Demandas de Projetos v2" (correções de Gates e catálogo de artefatos). */

const ARTEFATOS = {
  'SGP-01': { nome: 'Solicitação de Demanda', codigo: 'SGP-01', arquivo: '01- solicitacao-demanda.html' },
  'SGP-02': { nome: 'Canvas de Projeto', codigo: 'SGP-02', arquivo: '02 - canvas-de-projeto.html' },
  'SGP-03': { nome: 'TAP — Termo de Abertura do Projeto', codigo: 'SGP-03', arquivo: '03 - tap-termo-abertura-projeto.html' },
  'SGP-04': { nome: 'Planejamento e Desenvolvimento de Projeto', codigo: 'SGP-04', arquivo: '04 - planejamento-desenvolvimento-projeto.html' },
  'EAP':         { nome: 'EAP — Estrutura Analítica do Projeto', codigo: '— (sem código)', arquivo: '05 - eap-estrutura-analitica-projeto.html' },
  'SGP-06': { nome: 'SMP — Solicitação de Mudança do Projeto', codigo: 'SGP-06', arquivo: '06 - smp-solicitacao-mudanca-projeto.html' },
  'SGP-07': { nome: 'Ata de Reunião', codigo: 'SGP-07', arquivo: '07 - ata-reuniao-FORALF00340.html' },
  'SGP-08': { nome: 'Plano de Comunicação de Projeto', codigo: 'SGP-08', arquivo: '08 - plano-comunicacao-projeto.html' },
  'SGP-09': { nome: 'TEP — Termo de Encerramento do Projeto', codigo: 'SGP-09', arquivo: '09 - tep-termo-encerramento-projeto.html' },
  'SGP-10': { nome: 'RLA — Registro de Lições Aprendidas', codigo: 'SGP-10', arquivo: '10 - rla-licoes-aprendidas-FORALF00342.html' },
  'SGP-11':    { nome: 'Relatório de Situação de Projetos', codigo: 'SGP-11', arquivo: '11 - relatorio-situacao-FORALF11.html', extra: true },
  'SGP-12':    { nome: 'Relatório de Entregas e Benefícios', codigo: 'SGP-12', arquivo: '12 - relatorio-entregas-FORALF12.html' },
};

const GATES = {
  gate1: {
    id: 'gate1',
    nome: 'Gate 1 — Triagem',
    quando: 'D01.4 / D01.5, logo após o registro da demanda',
    quemDecide: 'Gestor Responsável',
    sobreArtefato: 'Solicitação de Demanda (SGP-01)',
    efeito: 'Aprova ou rejeita a entrada da demanda no portfólio.',
    apos: 'D01.5',
    diretriz: 'D01',
  },
  gate2: {
    id: 'gate2',
    nome: 'Gate 2 — Pactuação',
    quando: 'Fim do D02 (Planejamento), antes do início do D03 (Execução)',
    quemDecide: 'Dono do Negócio, perante instâncias hierarquicamente superiores',
    sobreArtefato: 'Relatório de Entregas e Benefícios (D06.1 / D06.2)',
    efeito: '"Autorização institucional para que todo o trabalho previsto seja prosseguido" — é o gate mais crítico: decide se o projeto de fato começa a ser executado.',
    apos: 'D02.10',
    diretriz: 'D02',
  },
};

const DIRETRIZES = [
  {
    id: 'D01', titulo: 'Demanda', tipo: 'sequencial', ordem: 1,
    resumo: 'Ponto de entrada de todo projeto: recebe, avalia e prioriza a demanda antes de virar projeto.',
    oQueE: 'Ponto de entrada de todo e qualquer projeto na instituição: a etapa em que uma necessidade, ideia ou problema institucional é identificado, registrado, avaliado e priorizado antes de ser formalizado como projeto. É um mecanismo de triagem.',
    porQue: 'Nem toda demanda deve virar projeto. A D01 protege a organização contra desperdício de recursos, confere foco às equipes e promove transparência nas decisões — um projeto mal iniciado tende a gerar retrabalho e desgastar a credibilidade da gestão de projetos.',
    entrega: ['Uma demanda avaliada, priorizada e aprovada, pronta para avançar para o planejamento.', 'Ou rejeitada, com justificativa clara e comunicada ao solicitante.'],
    pessoas: ['Solicitante da demanda', 'Gerente de Projetos', 'Gestor responsável pela aprovação'],
    grupos: [{ titulo: null, estrategias: [
      { id: 'D01.1', titulo: 'Recebimento e Registro de Demandas', pontos: [
        'Toda demanda deve ser formalmente registrada — nada pode subsistir apenas na forma verbal ou em comunicações informais.',
        'Usa-se a Solicitação de Demanda (SGP-01): nome do projeto, solicitante, justificativa, objetivo, escopo preliminar, prazo, orçamento e partes interessadas.',
        'Preenchimento é responsabilidade do solicitante, com apoio do Gerente de Projetos quando necessário. Demandas não registradas não entram no fluxo de avaliação.',
      ], artefatos: ['SGP-01'] },
      { id: 'D01.2', titulo: 'Análise de Viabilidade Preliminar', pontos: [
        'O Gerente de Projetos conduz uma avaliação inicial não exaustiva, apenas para decidir se a demanda pode avançar.',
        'Analisa: existência de problema/oportunidade concreta, disponibilidade de recursos, razoabilidade do prazo e clareza do objetivo.',
        'Se a inviabilidade for manifesta, a demanda é rejeitada nesse momento, com justificativa documentada.',
      ], artefatos: [] },
      { id: 'D01.3', titulo: 'Avaliação de Alinhamento Estratégico', pontos: [
        'Viabilidade técnica não é suficiente: a demanda precisa estar alinhada aos objetivos e prioridades institucionais.',
        'Iniciativas viáveis mas sem aderência estratégica podem ser postergadas, redirecionadas ou descartadas — recursos são finitos.',
        'A análise também verifica sobreposição com outros projetos em andamento no portfólio.',
      ], artefatos: [] },
      { id: 'D01.4', titulo: 'Priorização de Demandas', pontos: [
        'A ordem de ingresso no portfólio não pode ser por ordem de chegada ou insistência do demandante — exige critérios objetivos e transparentes.',
        'Critérios recomendados: impacto esperado nos resultados institucionais, urgência/prazo, volume de recursos necessários e alinhamento estratégico.',
        'Decisão é do Gestor Responsável, com suporte técnico do Gerente de Projetos.',
      ], artefatos: [] },
      { id: 'D01.5', titulo: 'Comunicação da Decisão', pontos: [
        'Todo solicitante tem direito a uma resposta formal — aprovação ou rejeição — com justificativa clara.',
        'Comunicação é responsabilidade do Gerente de Projetos, feita por e-mail ou Microsoft Teams, com registro formal.',
        'Em caso de rejeição, indicar se a demanda pode ser reavaliada futuramente ou se há alternativas.',
      ], artefatos: [] },
      { id: 'D01.6', titulo: 'Elaboração do Canvas de Projeto', pontos: [
        'O Canvas é o primeiro documento estruturado de um projeto aprovado: reúne problema, objetivo, benefícios esperados, produto, parcerias, restrições, riscos e custos estimados numa visão única.',
        'Elaboração colaborativa, com solicitante, Gerente de Projetos e principais interessados — reduz risco de rejeição e baixo engajamento adiante.',
        'É instrumento de alinhamento, não de detalhamento: assegura entendimento comum antes do planejamento formal (D02).',
      ], artefatos: ['SGP-07', 'SGP-02'] },
    ]}],
  },
  {
    id: 'D02', titulo: 'Planejamento', tipo: 'sequencial', ordem: 2,
    resumo: 'Transforma a demanda aprovada num plano de trabalho concreto e executável.',
    oQueE: 'O momento em que uma demanda aprovada e estruturada no Canvas se converte em um plano de trabalho concreto: escopo delimitado, entregas mapeadas, cronograma, recursos e custos estimados. A ponte entre a ideia aprovada e a execução orientada por controle e método.',
    porQue: 'Projetos que suprimem ou fazem o planejamento de forma superficial reproduzem os mesmos problemas: prazos não cumpridos, escopo indefinido, retrabalho e equipes sem direcionamento. Planejar não é burocracia, é o que torna possível executar com velocidade e segurança.',
    entrega: ['Projeto completamente estruturado: TAP assinado, stakeholders identificados, requisitos documentados, escopo definido, EAP elaborada, cronograma detalhado, recursos alocados e custos estimados.'],
    pessoas: ['Gerente de Projetos', 'Dono do Negócio', 'Equipe do projeto', 'Gestor responsável pela aprovação do TAP'],
    grupos: [{ titulo: null, estrategias: [
      { id: 'D02.1', titulo: 'Desenvolvimento do TAP', pontos: [
        'O Termo de Abertura autoriza formalmente a existência do projeto — transforma o Canvas (alinhamento) em compromisso formal, assinado pelo Gerente de Projetos, Dono do Negócio e Gestor Responsável.',
        'Consolida: identificação, justificativa, objetivo, escopo preliminar, premissas, restrições, riscos iniciais, equipe, marcos, orçamento e critérios de sucesso.',
        'Um projeto sem TAP aprovado não deve avançar para execução.',
      ], artefatos: ['SGP-07', 'SGP-03'] },
      { id: 'D02.2', titulo: 'Identificação e Análise de Stakeholders', pontos: [
        'Stakeholders são todas as pessoas, grupos ou organizações que podem afetar ou ser afetados pelo projeto — identificá-los cedo antecipa resistências e engaja apoiadores.',
        'A análise vai além de listar: mapeia grau de influência, nível de interesse e postura de cada parte.',
        'É atividade contínua, revisada ao longo de todo o projeto, não apenas uma vez.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D02.3', titulo: 'Coleta e Documentação de Requisitos', pontos: [
        'Requisitos são as condições que o projeto deve atender para satisfazer as partes interessadas — evita que o projeto seja construído sobre suposições.',
        'Coleta feita por entrevistas, reuniões e consultas diretas aos stakeholders identificados.',
        'Requisitos mal definidos tendem a ressurgir como mudanças de escopo durante a execução, com impacto em prazo e custo.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D02.4', titulo: 'Definição e Detalhamento do Escopo', pontos: [
        'Com os requisitos coletados, define-se com precisão o que está incluído no projeto e, igualmente importante, o que não está.',
        'Escopo claro previne crescimento não controlado por acréscimo de demandas não planejadas.',
        'Tudo que não estiver no escopo aprovado e for solicitado depois passa pelo processo formal de mudança (D04).',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D02.5', titulo: 'Definição e Sequenciamento de Atividades', pontos: [
        'As entregas do escopo são decompostas em atividades concretas, sequenciadas conforme suas dependências.',
        'Cada atividade tem responsável, duração estimada e dependências mapeadas — insumo direto para o cronograma.',
        'Categorização: fase, marco, atividade e tarefa (detalhamento opcional).',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D02.6', titulo: 'Estimativa de Recursos e Duração', pontos: [
        'Com atividades definidas, estimam-se recursos necessários (pessoas, equipamentos, materiais, serviços) e a duração de cada uma.',
        'Estimativa deve considerar a disponibilidade real da equipe, não a disponibilidade ideal — uma das etapas que mais exige honestidade do Gerente de Projetos.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D02.7', titulo: 'Desenvolvimento do Cronograma', pontos: [
        'Consolida, para cada atividade: responsável, data de início prevista, data de término prevista e dependências.',
        'É o produto direto das estratégias anteriores (atividades, recursos e durações já estimados).',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D02.8', titulo: 'Estimativa de Custos e Orçamento', pontos: [
        'A partir de atividades, recursos e cronograma, estima-se o custo de cada atividade e consolida-se o orçamento total do projeto.',
        'Serve de referência para o controle de custos ao longo da execução (D04).',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D02.9', titulo: 'Desenvolvimento da EAP', pontos: [
        'A Estrutura Analítica do Projeto organiza hierarquicamente todas as entregas e pacotes de trabalho do projeto.',
        'Funciona como mapa visual do escopo total, facilitando o controle e a comunicação sobre o que precisa ser entregue.',
      ], artefatos: ['SGP-07', 'EAP'] },
      { id: 'D02.10', titulo: 'Consolidar e Comunicar o Plano de Gerenciamento', pontos: [
        'Todo o planejamento (escopo, cronograma, custos, recursos, riscos, comunicação) é consolidado num plano único.',
        'O plano é comunicado a todos os envolvidos, garantindo que o projeto inicie com alinhamento e comprometimento institucional.',
        'É sobre este momento que incide o Gate 2 — Pactuação, antes do início da execução (D03).',
      ], artefatos: ['SGP-07', 'SGP-04'] },
    ]}],
  },
  {
    id: 'D03', titulo: 'Execução', tipo: 'sequencial', ordem: 3,
    resumo: 'O planejamento se transforma em ação: entregas concretas, equipe mobilizada.',
    oQueE: 'O momento em que o planejamento se transforma em ação. Os recursos alocados começam a produzir entregas concretas, as equipes são mobilizadas, as atividades são realizadas e o projeto passa a existir de fato. Fase de maior intensidade operacional do ciclo de vida.',
    porQue: 'Um projeto bem planejado pode fracassar por uma execução mal conduzida. É durante a execução que surgem imprevistos, conflitos de prioridade e desvios do que foi acordado. Gerenciar com método garante que o trabalho seja realizado conforme planejado e que as pessoas permaneçam engajadas.',
    entrega: ['As entregas do projeto realizadas conforme escopo, cronograma e orçamento aprovados.', 'Uma equipe engajada, com responsabilidades claras e comunicação estruturada.'],
    pessoas: ['Gerente de Projetos', 'Dono do Negócio', 'Equipe do projeto', 'Parceiros e fornecedores envolvidos na execução'],
    grupos: [{ titulo: null, estrategias: [
      { id: 'D03.1', titulo: 'Orientação e Gerenciamento do Trabalho', pontos: [
        'Dirigir e coordenar todas as atividades previstas no plano, assegurando execução dentro do escopo, prazo e recursos alocados — é a estratégia central da execução.',
        'O Gerente de Projetos deve manter proximidade constante com a equipe e identificar desvios antes que virem entraves. A execução não se gerencia à distância.',
        'Toda interação formal de acompanhamento deve ser registrada, garantindo rastreabilidade.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D03.2', titulo: 'Gerenciamento do Conhecimento do Projeto', pontos: [
        'Todo projeto gera conhecimento — sobre processos, organização, pessoas e soluções — que precisa ser capturado, organizado e disponibilizado para uso futuro.',
        'O Gerente de Projetos deve estimular o registro contínuo de decisões e lições aprendidas parciais ao longo da execução, alimentando a base de conhecimento da instituição.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D03.3', titulo: 'Gerenciamento da Qualidade', pontos: [
        'Verificar continuamente se as entregas atendem aos requisitos definidos no planejamento — qualidade não se avalia só no final.',
        'O Gerente de Projetos deve estabelecer critérios claros de aceite para cada entrega antes de iniciar o trabalho. Entregas fora do critério não devem ser aceitas, independentemente da pressão de prazo.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D03.4', titulo: 'Gerenciamento da Equipe', pontos: [
        'A equipe é o principal ativo do projeto: acompanhar desempenho individual e coletivo, identificar necessidades de desenvolvimento e resolver conflitos com agilidade.',
        'Garantir que as pessoas tenham recursos e informações necessárias, e reconhecer contribuições ao longo do processo — não só ao final.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D03.5', titulo: 'Gerenciamento das Comunicações', pontos: [
        'Garantir que as informações certas cheguem às pessoas certas, no momento adequado, pelos canais apropriados — conforme o Plano de Comunicação elaborado no planejamento.',
        'Falhas de comunicação geram retrabalho, conflitos e perda de engajamento. Deve ser prática regular, não apenas reativa a problemas.',
      ], artefatos: ['SGP-07', 'SGP-08'] },
      { id: 'D03.6', titulo: 'Implementação das Respostas aos Riscos', pontos: [
        'Riscos identificados no planejamento devem ser monitorados ativamente. Quando um risco se materializa, a resposta prevista deve ser implementada sem demora.',
        'Novos riscos que surgem na execução devem ser identificados, avaliados e registrados imediatamente — gestão de riscos é atividade contínua.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D03.7', titulo: 'Condução das Aquisições', pontos: [
        'Quando o projeto envolve contratação de fornecedores, serviços ou materiais externos, o processo deve ser conduzido de forma estruturada, dentro de prazos e valores orçados.',
        'Inclui acompanhamento de contratos em andamento e registro formal de qualquer desvio — aquisições mal geridas são causa comum de atraso e estouro de orçamento.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D03.8', titulo: 'Gerenciamento do Engajamento dos Stakeholders', pontos: [
        'Manter o engajamento ao longo da execução é tão importante quanto planejar bem: acompanhar continuamente o nível de comprometimento e agir proativamente diante de sinais de desmotivação ou resistência.',
        'Constrói-se com comunicação clara, reconhecimento de contribuições e participação nas decisões relevantes.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
    ]}],
  },
  {
    id: 'D04', titulo: 'Controle', tipo: 'sequencial', ordem: 4,
    resumo: 'Monitora continuamente a execução, identifica desvios e formaliza mudanças.',
    oQueE: 'O conjunto de estratégias responsáveis por monitorar e controlar tudo o que acontece durante a execução, verificando se o trabalho está conforme planejado, identificando desvios com antecedência e tomando ações corretivas. Ocorre em paralelo à execução, do início ao encerramento — não é uma etapa posterior.',
    porQue: 'Sem controle, mesmo o melhor planejamento se deteriora. Quanto mais tarde os desvios de escopo, prazo, custo ou qualidade são identificados, maior o impacto e o custo para corrigi-los. Controlar bem é o que diferencia projetos que entregam o que prometeram daqueles que chegam ao final com surpresas.',
    entrega: ['Visibilidade permanente sobre o desempenho do projeto em relação a escopo, cronograma, custo e qualidade aprovados.', 'Decisões de mudança formalizadas, desvios corrigidos e stakeholders informados.'],
    pessoas: ['Gerente de Projetos', 'Dono do Negócio', 'Equipe do projeto', 'Gestor responsável pela aprovação de mudanças'],
    grupos: [{ titulo: null, estrategias: [
      { id: 'D04.1', titulo: 'Monitoramento e Controle do Trabalho', pontos: [
        'Acompanhar, revisar e registrar o progresso do projeto em relação aos objetivos do plano de gerenciamento, dando visibilidade para agir antes que desvios se agravem.',
        'Deve ser contínuo, orientado por dados concretos: % de execução física dos marcos, situação das entregas, consumo de recursos e indicadores.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D04.2', titulo: 'Análise de Desempenho', pontos: [
        'Confronta o que foi realizado com o que foi planejado, avaliando se o projeto cumpre seus objetivos de prazo, custo, escopo e qualidade — insumo principal para decisões de correção de rumo.',
        'Deve ser conduzida com regularidade, com resultados comunicados com transparência, incluindo pontos de destaque e de atenção.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D04.3', titulo: 'Elaboração de Relatórios', pontos: [
        'Os relatórios de situação são o principal instrumento de comunicação do desempenho do projeto às partes interessadas — devem ser objetivos e periódicos.',
        'Um relatório deve ser instrumento de decisão, não apenas de registro. Um relatório que ninguém lê não cumpre sua função.',
      ], artefatos: ['SGP-07', 'SGP-04', 'SGP-11'] },
      { id: 'D04.4', titulo: 'Controle Integrado de Mudanças', pontos: [
        'Mudanças são inevitáveis; o que diferencia projetos bem gerenciados é como elas são tratadas — de forma formal, estruturada e com avaliação de impacto antes de aprovadas ou rejeitadas.',
        'Toda solicitação de mudança é registrada via SMP (Solicitação de Mudança do Projeto), com descrição, justificativa e análise de impacto. Aprovação é do Gestor Responsável, com subsídio técnico do Gerente de Projetos.',
        'Mudanças não formalizadas não devem ser implementadas.',
      ], artefatos: ['SGP-07', 'SGP-04', 'SGP-06'] },
      { id: 'D04.5', titulo: 'Validação do Escopo', pontos: [
        'Aceitação formal das entregas concluídas pelo Dono do Negócio — não basta estar tecnicamente pronta, precisa ser formalmente aceita.',
        'Deve ser conduzida de forma sistemática ao longo da execução, não concentrada no final, o que aumenta o risco de rejeições tardias e retrabalho caro.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D04.6', titulo: 'Controle de Escopo, Cronograma, Custos, Qualidade e Recursos', pontos: [
        'Controle integrado das principais dimensões do projeto: qualquer desvio deve ser identificado, registrado e tratado a tempo. As dimensões estão interligadas — um desvio em uma tende a impactar as demais.',
        'Erro comum a evitar: controlar apenas o cronograma enquanto custo e qualidade se deterioram silenciosamente.',
      ], artefatos: ['SGP-07', 'SGP-04', 'SGP-06'] },
      { id: 'D04.7', titulo: 'Monitoramento das Comunicações, Riscos, Aquisições e Demais Aspectos', pontos: [
        'Garante que dimensões frequentemente negligenciadas na execução permaneçam sob controle: comunicações verificadas quanto à efetividade, riscos reavaliados periodicamente, aquisições acompanhadas quanto a prazo, custo e qualidade.',
        'Problemas não identificados nessas áreas tendem a virar entraves nas fases finais, quando o custo de correção é maior.',
      ], artefatos: ['SGP-07', 'SGP-04', 'SGP-08', 'SGP-06'] },
    ]}],
  },
  {
    id: 'D05', titulo: 'Encerramento', tipo: 'sequencial', ordem: 5,
    resumo: 'Fecha formalmente o projeto e mantém viva a metodologia institucional.',
    oQueE: 'O conjunto de estratégias que conduzem o projeto ao fechamento formal e responsável: entregas aceitas, contratos encerrados, recursos desmobilizados, conhecimento registrado e benefícios comunicados. Não é um momento burocrático, mas o momento em que o projeto presta contas do que prometeu.',
    porQue: 'Projetos não formalmente encerrados deixam pontas soltas — contratos abertos, documentação dispersa, equipes sem desligamento adequado, benefícios nunca comunicados — o que compromete a credibilidade da gestão de projetos e impede o aprendizado organizacional.',
    entrega: ['Um projeto formalmente encerrado, com entregas aceitas, contratos encerrados, documentação arquivada, lições aprendidas registradas e benefícios comunicados.', 'Uma metodologia de gestão de projetos continuamente aprimorada.'],
    pessoas: ['Gerente de Projetos', 'Dono do Negócio', 'Equipe do projeto', 'Gestor responsável pela aceitação final', 'Todas as gerências envolvidas na melhoria contínua da metodologia'],
    grupos: [
      { id: 'D05.1', titulo: 'Finalizar e Gerir Conhecimento', estrategias: [
        { id: 'D05.1.1', titulo: 'Obtenção da Aceitação Final', pontos: [
          'Reconhecimento formal, pelo Dono do Negócio e Gestor Responsável, de que todas as entregas foram realizadas conforme os critérios de aceite. Sem essa formalização, o projeto permanece tecnicamente aberto.',
          'Divergências identificadas nessa etapa devem ser tratadas antes do encerramento definitivo.',
        ], artefatos: ['SGP-07', 'SGP-04'] },
        { id: 'D05.1.2', titulo: 'Encerramento de Contratos e Aquisições', pontos: [
          'Todos os contratos e aquisições vinculados ao projeto devem ser formalmente encerrados: cumprimento das obrigações, liquidação de pendências financeiras e recebimento definitivo dos produtos/serviços.',
          'Contratos não encerrados formalmente representam passivos financeiros, jurídicos ou operacionais para a organização.',
        ], artefatos: ['SGP-07', 'SGP-04'] },
        { id: 'D05.1.3', titulo: 'Encerramento Financeiro e Administrativo', pontos: [
          'Liquidação de obrigações financeiras pendentes, reconciliação dos custos realizados com o orçamento aprovado e encerramento de centros de custo.',
          'Essencial para que a organização compare o investimento realizado com os benefícios alcançados e alimente estimativas de projetos futuros.',
        ], artefatos: ['SGP-07', 'SGP-04'] },
        { id: 'D05.1.4', titulo: 'Desmobilização da Equipe e Recursos', pontos: [
          'Deve ser conduzida de forma ordenada e respeitosa: cada membro formalmente desvinculado, recursos liberados, contribuições reconhecidas.',
          'Uma desmobilização bem conduzida preserva o engajamento das pessoas para projetos futuros e evita ambiguidade sobre responsabilidades remanescentes.',
        ], artefatos: ['SGP-07', 'SGP-09'] },
        { id: 'D05.1.5', titulo: 'Consolidação e Arquivamento da Documentação', pontos: [
          'Toda a documentação gerada, da Solicitação de Demanda aos relatórios finais, deve ser consolidada e arquivada de forma acessível — registro histórico da iniciativa.',
          'Documentação dispersa ou incompleta compromete a rastreabilidade e a capacidade de aprendizado da organização.',
        ], artefatos: ['SGP-07', 'SGP-04', 'SGP-09'] },
        { id: 'D05.1.6', titulo: 'Condução de Lições Aprendidas', pontos: [
          'Principal mecanismo de aprendizado organizacional gerado pelos projetos. Deve ocorrer ao final e, preferencialmente, também em momentos intermediários.',
          'Conduzida de forma aberta e construtiva, sem julgamento. Lições não registradas e compartilhadas não geram aprendizado — ficam só na memória de quem participou.',
        ], artefatos: ['SGP-07', 'SGP-10'] },
        { id: 'D05.1.7', titulo: 'Elaboração e Distribuição do Relatório Final', pontos: [
          'Consolida resultados alcançados, benefícios entregues, desempenho em relação a escopo/cronograma/orçamento, decisões relevantes e lições aprendidas — é a prestação de contas do projeto.',
          'Distribuição deve alcançar todas as partes interessadas relevantes, não só quem participou da execução. Um projeto sem relatório final é um projeto que não prestou contas.',
        ], artefatos: ['SGP-07', 'SGP-09', 'SGP-10', 'SGP-04'] },
        { id: 'D05.1.8', titulo: 'Atualização da Base de Conhecimento', pontos: [
          'Metodologias aplicadas, soluções encontradas, boas práticas e lições aprendidas devem ser incorporadas à base de conhecimento da Gerência de Projetos.',
          'É o que transforma a experiência individual de cada projeto em patrimônio coletivo da organização.',
        ], artefatos: ['SGP-07', 'SGP-10'] },
        { id: 'D05.1.9', titulo: 'Emissão do Termo de Encerramento', pontos: [
          'O TEP formaliza oficialmente o encerramento da iniciativa, seja por conclusão bem-sucedida ou cancelamento justificado — libera formalmente recursos e responsabilidades.',
          'Assinado pelo Gerente de Projetos, Dono do Negócio e Gestor Responsável. Sem TEP emitido, o projeto permanece formalmente aberto.',
        ], artefatos: ['SGP-07', 'SGP-09'] },
        { id: 'D05.1.10', titulo: 'Comunicação das Entregas e Benefícios', pontos: [
          'Momento de comunicar amplamente os resultados alcançados a toda a organização e demais partes interessadas — ato de transparência e reconhecimento do esforço coletivo.',
          'Benefícios que não são comunicados são benefícios que não existem na percepção da organização.',
        ], artefatos: ['SGP-07', 'SGP-09', 'SGP-08'] },
      ]},
      { id: 'D05.2', titulo: 'Gerir Metodologia e Melhoria', estrategias: [
        { id: 'D05.2.1', titulo: 'Definição e Manutenção da Metodologia', pontos: [
          'A metodologia da instituição não é um documento estático: precisa ser continuamente revisada à luz da experiência acumulada. A Gerência de Projetos é responsável por mantê-la atualizada e relevante.',
          'Inclui revisão periódica das diretrizes e incorporação de melhorias identificadas nas lições aprendidas.',
        ], artefatos: ['SGP-07'] },
        { id: 'D05.2.2', titulo: 'Desenvolvimento e Manutenção de Templates', pontos: [
          'Templates, formulários e artefatos padronizados são os instrumentos que tornam a metodologia aplicável no dia a dia. Devem ser revisados periodicamente, orientados pela praticidade.',
          'Templates desatualizados ou de difícil uso comprometem a adesão à metodologia.',
        ], artefatos: ['SGP-07'] },
        { id: 'D05.2.3', titulo: 'Definição de Métricas e KPIs', pontos: [
          'Instrumentos para avaliar objetivamente o desempenho da gestão de projetos, não apenas de projetos individuais, mas do portfólio como um todo — foco em prazo, custo, qualidade e entrega de valor.',
          'Conjunto enxuto de indicadores mensuráveis, revisados periodicamente para continuar refletindo as prioridades estratégicas.',
        ], artefatos: ['SGP-07'] },
        { id: 'D05.2.4', titulo: 'Auditorias de Processos', pontos: [
          'Verificam se os projetos em andamento ou encerrados seguem as diretrizes e a metodologia estabelecidas — instrumento de qualidade, não de punição, com postura construtiva.',
          'Resultados alimentam diretamente a revisão da metodologia, atualização de templates e identificação de necessidades de capacitação.',
        ], artefatos: ['SGP-07'] },
        { id: 'D05.2.5', titulo: 'Gestão da Base de Conhecimento', pontos: [
          'Repositório central de lições aprendidas, boas práticas, templates, relatórios finais e referências metodológicas — precisa ser alimentada, organizada e efetivamente utilizada.',
          'Uma base que ninguém consulta não gera valor, é apenas um arquivo.',
        ], artefatos: ['SGP-07', 'SGP-10'] },
        { id: 'D05.2.6', titulo: 'Treinamento e Desenvolvimento da Equipe', pontos: [
          'A qualidade da gestão de projetos depende diretamente da capacidade das pessoas que a praticam — ações contínuas de treinamento, técnicas e comportamentais.',
          'Planejado com base nas lacunas identificadas em auditorias, lições aprendidas e evolução da metodologia. Não é evento pontual.',
        ], artefatos: ['SGP-07'] },
        { id: 'D05.2.7', titulo: 'Análise de Desempenho do Portfólio', pontos: [
          'Visão consolidada de todos os projetos em andamento e encerrados, avaliados em conjunto quanto a prazo, custo, qualidade e entrega de valor — indispensável para decisões estratégicas de priorização e alocação de recursos.',
          'Conduzida com periodicidade definida, apresentada aos gestores responsáveis de forma clara.',
        ], artefatos: ['SGP-07', 'SGP-04'] },
        { id: 'D05.2.8', titulo: 'Identificação e Implementação de Melhorias', pontos: [
          'Resultado natural da análise de portfólio, auditorias, lições aprendidas e uso cotidiano da metodologia — cada oportunidade deve ser avaliada, priorizada e implementada com responsável e prazo.',
          'Melhorias não implementadas são oportunidades perdidas.',
        ], artefatos: ['SGP-07'] },
        { id: 'D05.2.9', titulo: 'Comunicação das Ações', pontos: [
          'Revisões de processos, atualizações de templates, novos treinamentos, resultados de auditorias e melhorias implementadas devem ser comunicados de forma clara a todos os envolvidos.',
          'Uma metodologia que evolui em silêncio não evolui na prática.',
        ], artefatos: ['SGP-07'] },
      ]},
    ],
  },
  {
    id: 'D06', titulo: 'Governança e Tomada de Decisão', tipo: 'transversal', ordem: 6,
    resumo: 'Transversal a D01–D05: garante que as decisões certas sejam tomadas pelas pessoas certas, no momento certo.',
    oQueE: 'Diretriz transversal que atravessa todas as demais fases do ciclo de vida, do D01 ao D05, sem se restringir a um momento específico. Define o direcionamento estratégico da gestão de projetos e busca conquistar e preservar a confiança das partes interessadas. Concentra-se em quem toma as decisões (papéis e autoridade) e em como elas são tomadas (processos para mudanças, entraves e encaminhamentos).',
    porQue: 'Projetos sem governança tendem a acumular decisões postergadas, entraves não resolvidos e mudanças não formalizadas, até que o peso dessas pendências comprometa irreversivelmente o resultado. A governança não é controle pelo controle — é o mecanismo que protege o projeto e a organização.',
    entrega: ['Um modelo de governança estruturado, com papéis definidos, reuniões periódicas e fluxo de informações claro entre todos os níveis de decisão.', 'Entraves resolvidos, encaminhamentos monitorados e resultados reportados com transparência e regularidade.'],
    pessoas: ['Gerente de Projetos', 'Dono do Negócio', 'Gestor Responsável', 'Alta Gestão', 'Todas as partes interessadas com papel de decisão nos projetos'],
    grupos: [{ titulo: null, estrategias: [
      { id: 'D06.1', titulo: 'Elaboração do Relatório de Entregas e Benefícios', pontos: [
        'Documento que consolida o planejamento do projeto numa visão executiva: responsáveis, justificativa, objetivo, alinhamento estratégico, indicadores, previsão financeira, parceiros e entregas previstas.',
        'Elaborado ao final do planejamento, antes do início da execução, pelo Gerente de Projetos em conjunto com o Dono do Negócio.',
      ], artefatos: ['SGP-07', 'SGP-04', 'SGP-12'] },
      { id: 'D06.2', titulo: 'Pactuação do Relatório de Entregas e Benefícios', pontos: [
        'Após elaborado, o relatório deve ser formalmente pactuado: os compromissos nele contidos são assumidos pelo Dono do Negócio perante as instâncias hierarquicamente superiores.',
        'Este é o momento do Gate 2 — a autorização institucional para que todo o trabalho previsto seja prosseguido.',
      ], artefatos: ['SGP-07', 'SGP-04', 'SGP-12'] },
      { id: 'D06.3', titulo: 'Gestão do Relatório de Entregas e Benefícios', pontos: [
        'Uma vez pactuado, o relatório é um instrumento vivo, gerido continuamente ao longo da execução, sensível a mudanças de diretrizes, novos cenários e decisões das reuniões de governança.',
        'Um relatório desatualizado perde sua função como instrumento de governança.',
      ], artefatos: ['SGP-07', 'SGP-04', 'SGP-12'] },
      { id: 'D06.4', titulo: 'Realização de Reuniões de Governança', pontos: [
        'Mecanismo central pelo qual informações, decisões e encaminhamentos percorrem todos os níveis hierárquicos — do operacional ao estratégico. Reuniões estruturadas, com pauta definida, foco em decisão, não em relato.',
        'Recomenda-se organizar por níveis: reuniões de equipe mais frequentes, gerenciais quinzenais e estratégicas mensais. Cada nível resolve o que está em sua governabilidade e encaminha ao nível superior só o que exige decisão maior.',
      ], artefatos: ['SGP-07', 'SGP-04', 'SGP-12'] },
      { id: 'D06.5', titulo: 'Reporte de Resultados', pontos: [
        'Momento em que o Gerente de Projetos e o Dono do Negócio apresentam às instâncias de decisão o desempenho do projeto — não apenas o que foi feito, mas o que foi alcançado em indicadores, entregas e benefícios.',
        'Reporte objetivo e orientado para decisão: destaca entraves que não puderam ser resolvidos em níveis anteriores e pontos que requerem deliberação superior.',
      ], artefatos: ['SGP-07', 'SGP-04', 'SGP-11'] },
      { id: 'D06.6', titulo: 'Gestão de Entraves', pontos: [
        'Entraves são obstáculos que, se não tratados com agilidade, comprometem o alcance dos resultados. Todo entrave identificado deve ser registrado imediatamente, classificado e encaminhado para resolução na instância competente.',
        'Três situações de acompanhamento: Aberto, Resolvido e Cancelado. A relação entre entraves abertos e resolvidos é um dos principais indicadores da saúde da governança do projeto.',
      ], artefatos: ['SGP-07', 'SGP-04', 'SGP-11'] },
      { id: 'D06.7', titulo: 'Gestão de Encaminhamentos', pontos: [
        'Encaminhamentos são deliberações tomadas para resolver entraves ou avançar em aspectos que requerem ação específica. Todo encaminhamento precisa de um único responsável, prazo definido e situação acompanhada nas reuniões seguintes.',
        'Três situações: Aberto, Resolvido e Cancelado. Encaminhamentos sem responsável ou prazo não devem ser aceitos — a indefinição é inimiga da governança.',
      ], artefatos: ['SGP-07', 'SGP-04', 'SGP-11'] },
    ]}],
  },
  {
    id: 'D07', titulo: 'Liderança, Equipes e Cultura', tipo: 'transversal', ordem: 7,
    resumo: 'Transversal a D01–D05: a dimensão humana — identificar, engajar, desenvolver e liderar pessoas.',
    oQueE: 'Diretriz transversal, presente em todas as fases do ciclo de vida, do D01 ao D05. Reconhece que projetos são realizados por pessoas e que o sucesso de qualquer iniciativa depende da capacidade de identificar, engajar, desenvolver e liderar as pessoas envolvidas. Trata da dimensão humana que nenhuma ferramenta ou documento substitui.',
    porQue: 'Um dos maiores riscos em qualquer projeto não está no escopo, cronograma ou orçamento, mas nas pessoas. Pessoas desengajadas não entregam; pessoas sem desenvolvimento não evoluem; e uma organização sem cultura de projetos repete os mesmos erros, independentemente da qualidade de seus processos.',
    entrega: ['Uma rede de pessoas envolvidas nos projetos identificada, mapeada e gerida de forma estruturada.', 'Pessoas engajadas, comprometidas e continuamente desenvolvidas.', 'Uma cultura organizacional orientada para o pensamento por projetos e para a entrega de valor.'],
    pessoas: ['Gerente de Projetos', 'Dono do Negócio', 'Equipe do projeto', 'Todas as partes interessadas que impactam ou são impactadas pelos projetos da instituição'],
    grupos: [{ titulo: null, estrategias: [
      { id: 'D07.1', titulo: 'Criação da Rede de Pessoas Envolvidas nos Projetos', pontos: [
        'A rede reúne todos os indivíduos que, de alguma forma, impactam ou são impactados por um projeto, positiva ou negativamente, direta ou indiretamente — identificá-los desde o início antecipa resistências e mobiliza apoiadores.',
        'Pode ser extensa e heterogênea: solicitantes, gestores, equipe, parceiros e fornecedores. Toda nova pessoa identificada deve ser incorporada e registrada.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D07.2', titulo: 'Gestão da Rede de Pessoas Envolvidas nos Projetos', pontos: [
        'Identificar é só o primeiro passo: a rede precisa ser gerida continuamente, promovendo participação, coletando feedbacks e reduzindo riscos relacionados a pessoas.',
        'Recomenda-se analisar grau de influência, nível de interesse e postura de cada parte — de empreendedores que adotam o projeto como próprio até fortes opositores que requerem ação mais direta.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D07.3', titulo: 'Engajamento das Pessoas Envolvidas nos Projetos', pontos: [
        'Engajamento é o estado em que as pessoas estão genuinamente comprometidas, contribuindo ativamente para o sucesso — uma das responsabilidades centrais do Gerente de Projetos.',
        'Promovido por propósito, emoções positivas e redução de barreiras (custo, incerteza, capacidade). Alavancas: apelo social, incentivos, linguagem direta, simplificação e experimentação.',
      ], artefatos: ['SGP-07', 'SGP-04'] },
      { id: 'D07.4', titulo: 'Orientação e Realização de Capacitações', pontos: [
        'A qualidade da gestão de projetos depende diretamente da capacidade das pessoas que a praticam — não basta ter processos e ferramentas, é preciso saber usá-los.',
        'Capacitações planejadas com base em lacunas reais, cobrindo competências técnicas (planejamento, controle de mudanças, riscos) e comportamentais (liderança, comunicação, negociação, resiliência, empatia).',
      ], artefatos: ['SGP-07', 'SGP-04'] },
    ]}],
  },
];

const SEQUENCIAIS = DIRETRIZES.filter(d => d.tipo === 'sequencial');
const TRANSVERSAIS = DIRETRIZES.filter(d => d.tipo === 'transversal');

function getDiretriz(id) {
  return DIRETRIZES.find(d => d.id === id);
}
