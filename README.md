# auto-fechamento
Uma forma de automatizar o fechamento geral das lojas TACO, garantindo velocidade e precisão diária para os operadores de caixa. Aplicativo de uso interno.

![Rel. Fechamento caixa](https://github.com/user-attachments/assets/bcba2385-5dbd-4182-9405-40b17a68977e)

![Saldo em caixa](https://github.com/user-attachments/assets/a50442cc-0d8f-4aa6-825f-078e7d9b2da0)

Ambas as abas serão automaticamente preenchidas com os valores retirados do relatório do Linx Manager.

## Como executar o projeto

1. Clone o repositório:
    ```bash
    git clone https://github.com/PedroLucasSinuso/auto-fechamento.git
    ```

2. Navegue até o diretório do projeto:
    ```bash
    cd auto-fechamento
    ```

3. Instale as dependências:
    ```bash
    pip install -r requirements.txt
    ```

4. Execute o aplicativo:
    ```bash
    streamlit run frontend.py
    ```

## Acessar o projeto online

Você também pode acessar o projeto online através do seguinte link:
[Auto-fechamento TACO](https://auto-fechamento.streamlit.app)

## Como usar o aplicativo

1. Abra o Menu Relatórios do Linx Manager (requer senha administrativa):
   
   ![Menu Relatorio](https://github.com/user-attachments/assets/20d3fdb9-0b88-4379-b2cb-951a6df444c7)

2. Certifique-se que as configurações estejam de acordo com o dia a ser fechado e clique em Pesquisar(F3):
   
   ![Configurações Menu Relatório](https://github.com/user-attachments/assets/e5358cb2-d711-48d0-a22c-aec0e30bfccc)


3. Clique na guia "Relatórios" e depois em "Resumos do caixa":
   
   ![Resumos do Caixa](https://github.com/user-attachments/assets/341e36e7-ef5d-44d1-9b71-6d8e3501b7ab)

4. A visualização do relatório de vendas por terminal será exibida. Clique no ícone de exportar destacado:
   
   ![Visualizacao_Relatorio](https://github.com/user-attachments/assets/4a41cf85-209f-4e2b-a9eb-6687c1589aae)

5. Será exibida a caixa de diálogo de opções de exportação. Selecione a exportação "Word - flow layout" e, em "File name", clique no botão reticências. Na janela seguinte, nomeie o arquivo de exportação .doc e salve-o. Dica: Reserve uma pasta exclusiva para as exportações diárias do Linx. Além disso, sempre salve o arquivo com a data e o mês no formato dd-mm (ex: 22-03.doc)

   ![Exportacao_Correta](https://github.com/user-attachments/assets/660d8ac9-78cc-40e5-a174-e3f21f5743dd)

6. Abra o app do Autofechamento e faça upload do arquivo word na opção da esquerda e faça upload da planilha de fechamento do dia anterior:

    ![Autofechamento](https://github.com/user-attachments/assets/37766c6e-667d-4347-a94d-c588afa5c04a)

7. Caso seja dia de coleta da Prosegur, marque a opção de Coleta. Faça o mesmo para recebimento de Troco e adicione o valor enviado/recebido:

   ![Autofechamento](https://github.com/user-attachments/assets/55583e65-4b96-4894-affb-7e9214af3385)

8. Após atualizar o valor do Fundo de Caixa, clique em Gerar Relatório. O aplicativo avisará caso haja divergência:

   ![Autofechamento](https://github.com/user-attachments/assets/ee373ced-acfa-4160-a491-0195be8d9ec9)

   ![Autofechamento](https://github.com/user-attachments/assets/8656eaf9-4e66-488a-95b2-2615443e00b9)

9. Basta clicar em "Baixar planilha de fechamento" que o programa fará download dela preenchida conforme os valores do Linx:

Obs: Lembrar de SEMPRE abrir a planilha, clicar em "Habilitar Edição" e salvar a planilha na pasta de fechamento. Caso a planilha não seja aberta e salva, o excel não calculará os valores referentes ao saldo em espécie da loja e a planilha do dia anterior não será confiável.

   ![Autofechamento](https://github.com/user-attachments/assets/dfc4aa1e-5c01-4e89-80e2-53570afd1307)








