import win32com.client
import os
import time
import openai
import msvcrt

#API Key
openai.api_key = os.getenv("OPENAI_API_KEY")

# Generate text with OpenAI
try:
    def ai_response(user_prompt: str):
     
        response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
                {"role": "system", "content": "You are  Print.ai developped by Nouhan Kourouma."},
                {"role": "user", "content": "Who won the world series in 2020?"},
                {"role": "assistant", "content": "The Los Angeles Dodgers won the World Series in 2020."},
                {"role": "user", "content":user_prompt }
            ]
        ) 
             
        
        return response["choices"][0]["message"]["content"]  
except Exception as e:
    raise e


if __name__ == "__main__":
   while True:
        user_prompt = input("Your question:  ")
        RESPONSE = ai_response(user_prompt)


        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        if word.Documents.Count > 0:
            doc = word.Documents(1)
        else:
            doc = word.Documents.Add()
        doc.Range(doc.Content.End-1, doc.Content.End-1).Select()
        text = RESPONSE
        for char in text:
            word.Selection.TypeText(char)
            time.sleep(0.01)  
            
        if msvcrt.kbhit() and ord(msvcrt.getch()) == 27: 
            print("La boucle a été interrompue par l'utilisateur.")
            SystemExit(0)
