import streamlit as st
import openai
from openai import OpenAI
import os
from dotenv import load_dotenv
import pandas as pd

# Carregar variáveis de ambiente do .env
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")

# Cria o cliente OpenAI
client = OpenAI()


def process_excel_and_generate_response(prompt_input, uploaded_file):
    """Processa o arquivo Excel, gera a prompt e obtém a resposta do modelo."""
    try:
        # Tenta ler o arquivo Excel usando openpyxl (ou xlrd se openpyxl falhar)
        df = pd.read_excel(uploaded_file, engine='openpyxl', skiprows=0)

        if df.empty:
            st.error("O arquivo XLSX está vazio. Por favor, envie um arquivo com dados.")
            return

        # Define o tamanho do chunk (em linhas)
        chunk_size = 800

        # Converte o DataFrame em chunks de texto
        chunks_text = []
        for i in range(0, len(df), chunk_size):
            chunk = df[i:i + chunk_size]
            chunks_text.append(chunk.to_string(index=False))

        # Une os chunks de texto em uma única string
        combined_chunk_text = "\n\n".join(chunks_text)

        # Cria a prompt completa, incluindo o conteúdo dos chunks.
        full_prompt = f"""{prompt_input}

Conteúdo do arquivo XLSX:
{combined_chunk_text}
"""

         # Imprime o conteúdo do Excel antes da requisição (com melhor formatação)
        st.write("**Conteúdo do arquivo Excel (parte visível):**")
        st.dataframe(df)  # Mostre o DataFrame como uma tabela


        try:
            # Faz a solicitação à API do OpenAI, com tratamento de erros.
            response = client.chat.completions.create(
                model="gpt-4o-mini", 
                messages=[
                    {"role": "user", "content": full_prompt}
                ],
                temperature=0.7,
            )

            st.markdown("### Resposta do modelo:")
            st.markdown(response.choices[0].message.content.strip())

        except openai.error.RateLimitError as e:
            st.error(f"Limite de requisições atingido. Tente novamente mais tarde: {e}")
        except openai.error.APIError as e:
            st.error(f"Erro na API do OpenAI: {e}")
        except Exception as e:
            st.error(f"Erro ao processar a solicitação: {e}")

    except FileNotFoundError:
        st.error("Arquivo não encontrado. Certifique-se de que o arquivo Excel está no local correto.")
    except pd.errors.EmptyDataError:
        st.error("O arquivo Excel está vazio.")
    except Exception as e:
        st.error(f"Erro ao ler o arquivo XLSX: {e}")


# Interface do Streamlit
st.title("Modelo GPT com Upload de Arquivo XLSX")
st.write("Envie um arquivo XLSX e uma prompt para obter a resposta do modelo.")

prompt_input = st.text_area("Digite a prompt para o modelo:")
uploaded_file = st.file_uploader("Envie um arquivo XLSX", type=["xlsx"])

if st.button("Enviar para o modelo"):
    process_excel_and_generate_response(prompt_input, uploaded_file)