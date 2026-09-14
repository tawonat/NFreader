# Mapa semântico de serviços

O cadastro de serviços possui poucos códigos, então o sistema não usa apenas fuzzy matching. O `main.py` combina:

1. termos fortes e sinônimos do serviço;
2. termos de apoio;
3. termos de exclusão para desempate;
4. similaridade textual da descrição;
5. regras específicas para serviços parecidos, como assistência técnica x hora-homem e destinação de resíduo x locação de caçamba.

Exemplos de interpretação:

- lixo/resíduo/destinação/descarte -> 771206
- caçamba/entulho com caçamba -> 830132
- gerador/grupo gerador -> 900136
- datadora/codificador/marcador de lote -> 900076
- plataforma elevatória -> 780117
- martelete/rompedor -> 717703
- andaime/escoramento -> 715237
- concreto/concretagem -> 730170
- calibração/aferição/metrologia -> 700186
- hospedagem/hotel/diária -> 760000
- frete/transporte/carga -> 780007
- assistência/manutenção/reparo/fabricação/usinagem -> 750027
- horas técnicas/hora-homem/mão de obra -> 750685

Descrições genéricas como `Prestação de serviço` não recebem um código automaticamente e são enviadas para revisão.

Para mudar sinônimos ou acrescentar novas formas de descrição, edite `SERVICE_SEMANTIC_MAP` em `main.py`.


### NFS-e genérica
Quando a NFS-e declara apenas `Prestacao de servico`, o sistema preserva essa descrição na saída e usa, exclusivamente para classificação, o texto descritivo do bloco `SERVIÇO PRESTADO` quando disponível. Isso evita inventar uma descrição do prestador.
