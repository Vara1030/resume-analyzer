import streamlit as st
import os
from dotenv import load_dotenv
import PyPDF2
import docx
import re

load_dotenv()

st.set_page_config(
    page_title="Resume Analyzer",
    page_icon="📄",
    layout="wide"
)

st.title("📄 LLM-Powered Resume Analyzer")
st.markdown("Upload resumes and get AI-powered analysis")

# Sidebar for API key
with st.sidebar:
    st.header("⚙️ Configuration")
    api_key = st.text_input("OpenAI API Key", type="password")
    if api_key:
        os.environ["OPENAI_API_KEY"] = api_key
        st.success("✅ API Key set!")

# Main content
col1, col2 = st.columns(2)

with col1:
    st.subheader("📤 Upload Resume")
    uploaded_file = st.file_uploader(
        "Choose a resume file",
        type=['pdf', 'docx']
    )
    
    job_description = st.text_area(
        "📋 Job Description (Optional)",
        placeholder="Paste job description here...",
        height=150
    )

with col2:
    st.subheader("📊 Analysis Results")
    
    if uploaded_file:
        st.success("✅ File uploaded successfully!")
        
        # Extract text from file
        text = ""
        try:
            if uploaded_file.name.endswith('.pdf'):
                pdf_reader = PyPDF2.PdfReader(uploaded_file)
                for page in pdf_reader.pages:
                    text += page.extract_text()
            elif uploaded_file.name.endswith('.docx'):
                doc = docx.Document(uploaded_file)
                for para in doc.paragraphs:
                    text += para.text + "\n"
            
            # Simple skill extraction
            skills_list = [
                'python', 'java', 'javascript', 'react', 'angular', 'vue',
                'node.js', 'django', 'flask', 'spring', 'aws', 'azure', 'gcp',
                'docker', 'kubernetes', 'git', 'sql', 'mongodb', 'postgresql',
                'mysql', 'html', 'css', 'typescript', 'c++', 'machine learning',
                'deep learning', 'nlp', 'ai', 'data science', 'linux'
            ]
            
            found_skills = []
            text_lower = text.lower()
            for skill in skills_list:
                if skill in text_lower:
                    found_skills.append(skill)
            
            # Display extracted info
            st.markdown("### 📝 Extracted Information")
            
            col_skills, col_edu = st.columns(2)
            with col_skills:
                st.markdown("**Skills Found:**")
                for skill in found_skills[:10]:
                    st.markdown(f"- {skill}")
            
            with col_edu:
                st.markdown("**Education:**")
                # Simple education detection
                edu_keywords = ['b.tech', 'bachelor', 'master', 'degree', 'university', 'college']
                sentences = text.split('.')
                edu_found = []
                for sentence in sentences:
                    for keyword in edu_keywords:
                        if keyword in sentence.lower():
                            edu_found.append(sentence.strip())
                            break
                for edu in edu_found[:3]:
                    st.markdown(f"- {edu}")
            
            # If API key is provided, show AI analysis button
            if api_key or os.getenv("OPENAI_API_KEY"):
                if st.button("🚀 Generate AI Analysis"):
                    try:
                        import openai
                        client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
                        
                        prompt = f"""
                        Analyze this resume and provide:
                        1. Profile Summary (2-3 sentences)
                        2. Key Strengths
                        3. Areas for Improvement
                        4. Suggested Job Roles
                        
                        RESUME TEXT:
                        {text[:3000]}
                        """
                        
                        if job_description:
                            prompt += f"\n\nJOB DESCRIPTION:\n{job_description}"
                        
                        with st.spinner("🧠 AI is analyzing..."):
                            response = client.chat.completions.create(
                                model="gpt-3.5-turbo",
                                messages=[
                                    {"role": "system", "content": "You are an expert resume analyst."},
                                    {"role": "user", "content": prompt}
                                ],
                                temperature=0.7,
                                max_tokens=800
                            )
                            
                            st.markdown("### 🤖 AI Analysis")
                            st.write(response.choices[0].message.content)
                    except Exception as e:
                        st.error(f"AI Analysis Error: {str(e)}")
                        st.info("Make sure your OpenAI API key is correct.")
            else:
                st.info("💡 Add your OpenAI API key in the sidebar for AI analysis.")
                
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
    else:
        st.info("👆 Upload a resume to get started")
        
        st.markdown("### 📌 Quick Guide")
        st.markdown("""
        1. Upload a resume (PDF or DOCX)
        2. (Optional) Add a job description
        3. Add OpenAI API key in sidebar
        4. Click 'Generate AI Analysis'
        """)

st.markdown("---")
st.markdown("""
    <div style='text-align: center; color: gray;'>
        Built with ❤️ using Streamlit
    </div>
""", unsafe_allow_html=True)