from flask import Flask, render_template, request, session
import cohere
import os
from docx import Document

app = Flask(__name__)
app.secret_key = os.urandom(24)

co = cohere.Client("uY6A7W64ORC45Jy2RuVxURTnA02gmaO2BEQkOZOa")  

def read_docx_text(docx_path):
    try:
        doc = Document(docx_path)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        return f"[Error loading {docx_path}] {str(e)}"

dress_code_text = read_docx_text("Dress_Code.docx")
leave_policy_text = read_docx_text("Leave_Policy.docx")

def ask_cohere(question):
    try:
        response = co.chat(
            model="command-r-plus",
            message=question,
            documents=[
                {"title": "Dress Code", "text": dress_code_text},
                {"title": "Leave Policy", "text": leave_policy_text}
            ],
            temperature=0.5
        )
        return response.text.strip()
    except Exception as e:
        return f"[Error] {str(e)}"

@app.route("/", methods=["GET", "POST"])
def index():
    session.setdefault("chat_history", [])
    if request.method == "POST":
        question = request.form.get("question", "")
        answer = ask_cohere(question)
        session["chat_history"].append({"sender": "user", "text": question})
        session["chat_history"].append({"sender": "bot", "text": answer})
        session.modified = True
    return render_template("index.html", chat_history=session["chat_history"])

@app.route("/reset")
def reset():
    session.pop("chat_history", None)
    return "Chat history cleared."

if __name__ == "__main__":
    app.run(debug=True)
