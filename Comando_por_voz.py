import speech_recognition as sr
import os
import sys
import time
import traceback

# Variável do programa para controlar o estado do programa
executando = True

# Função para ouvir e reconhecer a fala
def ouvir_microfone():
    global executando
    # Habilita microfone do usuário
    r = sr.Recognizer()

    try:
        # Listar microfones disponíveis para diagnóstico
        print("\nMicrofones disponíveis: ")
        for i, mic_name in enumerate(sr.Microphone.list_microphone_names()):
            print(f"{i}: {mic_name}")

        
        # Usando o microfone
        with sr.Microphone() as source:
            print("\nAjustando para ruído ambiente...")
            r.adjust_for_ambient_noise(source, duration=2)    

            # Frase para usuario dizer algo
            print("Diga alguma coisa: (ou 'Fechar' para sair)")

            # Armazena o que foi dito numa variável
            audio = r.listen(source, timeout=5, phrase_time_limit=5)

            try:
                # Passa a variável para o algoritmo reconhecedor de padões
                frase = r.recognize_google(audio, language='pt-BR').lower()

                print(("\nVocê disse: ", frase))

                if "navegador" in frase:
                    os.system("start Chrome.exe")
                elif "Excel" in frase:
                    os.system("start Excel.exe")
                elif "PowerPoint" in frase:
                    os.system("start POWERPNT.exe")
                elif "Edge" in frase:
                    os.system("start msedge.exe")
                elif "Fechar" in frase or "sair" in frase or "encerrar" in frase:
                    print("Encerando programa...")
                    executando = False
                else:
                    print("Comando não reconhecido!")  

            except sr.UnknownValueError:
                print("Não entendi o audio!")
                return False
            except sr.RequestError as e:
                print(f"Erro no serviço de reconhecimento: {str(e)}")
        
    except OSError:
        print(f"Erro: Microfone não disponível ou não encontrado!: {str(e)}")
    except Exception as e:
        print(f"Erro inesperado: {str(e)}")

if __name__ == "__main__":
    print("Sistema de reconhecimento de voz iniciado...")
    print("Diga  'fechar' para encerrar programa!")

    while executando:
        try:
            ouvir_microfone()
            time.sleep(1)
        except KeyboardInterrupt:
            print("\nPrograma interrompido pelo usuário!")
            executando = False

    print("Programa encerrado com sucesso!")
    sys.exit(0)
