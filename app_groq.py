import streamlit as st
import os
import PyPDF2
import docx
from groq import Groq

# Page Config
st.set_page_config(
    page_title="Resume Analyzer",
    page_icon="📄",
    layout="wide"
)

# Custom CSS - Enhanced Readability
st.markdown("""
    <style>
        /* Import Google Fonts */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
        
        * {
            font-family: 'Inter', sans-serif;
        }
        
        /* Hide default Streamlit elements */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        
        /* Main background with gradient */
        .stApp {
            background: linear-gradient(135deg, #f5f7fa 0%, #e8ecf1 100%);
        }
        
        /* Premium Header */
        .header-container {
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
            padding: 2.5rem 3rem;
            border-radius: 20px;
            margin-bottom: 2rem;
            position: relative;
            overflow: hidden;
            box-shadow: 0 20px 60px rgba(0,0,0,0.15);
        }
        
        .header-container::before {
            content: '';
            position: absolute;
            top: -50%;
            right: -20%;
            width: 500px;
            height: 500px;
            background: rgba(255,255,255,0.03);
            border-radius: 50%;
        }
        
        .header-title {
            font-size: 3rem;
            font-weight: 800;
            color: white;
            margin: 0;
            letter-spacing: -0.02em;
            position: relative;
            z-index: 1;
        }
        
        .header-subtitle {
            font-size: 1.1rem;
            color: rgba(255,255,255,0.8);
            margin-top: 0.5rem;
            font-weight: 300;
            position: relative;
            z-index: 1;
        }
        
        .header-badge {
            display: inline-block;
            background: rgba(100, 200, 255, 0.2);
            backdrop-filter: blur(10px);
            padding: 0.3rem 1rem;
            border-radius: 100px;
            color: #64c8ff;
            font-size: 0.75rem;
            margin-top: 0.5rem;
            border: 1px solid rgba(100, 200, 255, 0.2);
            position: relative;
            z-index: 1;
        }
        
        /* Premium Cards */
        .premium-card {
            background: white;
            padding: 2rem;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.06);
            border: 1px solid rgba(0,0,0,0.04);
            height: 100%;
        }
        
        .card-title {
            font-size: 1.1rem;
            font-weight: 600;
            color: #1a1a2e;
            margin-bottom: 1.2rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        /* Upload Area */
        .upload-area {
            border: 2px dashed #d0d5dd;
            border-radius: 12px;
            padding: 2.5rem 1.5rem;
            text-align: center;
            background: #fafbfc;
            transition: all 0.3s ease;
        }
        
        .upload-area:hover {
            border-color: #4a90d9;
            background: #f0f4ff;
        }
        
        .upload-icon {
            font-size: 2.5rem;
            margin-bottom: 0.5rem;
        }
        
        /* Skill Tags */
        .skill-tag {
            display: inline-block;
            background: linear-gradient(135deg, #4a90d9 0%, #357abd 100%);
            color: white;
            padding: 0.25rem 0.8rem;
            border-radius: 100px;
            font-size: 0.7rem;
            font-weight: 500;
            margin: 0.15rem;
        }
        
        /* Info items */
        .info-item {
            background: #f8f9fa;
            padding: 0.6rem 1rem;
            border-radius: 8px;
            margin: 0.3rem 0;
            border-left: 3px solid #4a90d9;
            color: #1a1a2e;
        }
        
        /* Analysis Result - FIXED TEXT VISIBILITY */
        .analysis-result {
            background: #f8f9fa;
            padding: 1.5rem;
            border-radius: 12px;
            border-left: 4px solid #4a90d9;
            color: #1a1a2e !important;
            line-height: 1.8;
            font-size: 0.95rem;
        }
        
        .analysis-result p, 
        .analysis-result li, 
        .analysis-result div,
        .analysis-result span {
            color: #1a1a2e !important;
        }
        
        .analysis-result h1, 
        .analysis-result h2, 
        .analysis-result h3, 
        .analysis-result h4 {
            color: #0f3460 !important;
            margin-top: 1rem;
            margin-bottom: 0.5rem;
        }
        
        .analysis-result strong {
            color: #0f3460 !important;
        }
        
        /* Button */
        .premium-btn {
            background: linear-gradient(135deg, #4a90d9 0%, #357abd 100%);
            color: white;
            border: none;
            padding: 0.7rem 2rem;
            border-radius: 10px;
            font-weight: 600;
            font-size: 0.95rem;
            width: 100%;
            cursor: pointer;
            box-shadow: 0 4px 15px rgba(74, 144, 217, 0.3);
            transition: all 0.3s ease;
        }
        
        .premium-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 25px rgba(74, 144, 217, 0.4);
        }
        
        /* Sidebar */
        .css-1d391kg {
            background: rgba(255,255,255,0.9);
            backdrop-filter: blur(20px);
            border-right: 1px solid rgba(0,0,0,0.05);
        }
        
        /* Responsive */
        @media (max-width: 768px) {
            .header-container {
                padding: 1.5rem;
            }
            .header-title {
                font-size: 2rem;
            }
            .premium-card {
                padding: 1.2rem;
            }
        }
        
        .fade-in {
            animation: fadeInUp 0.6s ease forwards;
        }
        
        @keyframes fadeInUp {
            from {
                opacity: 0;
                transform: translateY(20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
    <div class="header-container fade-in">
        <div class="header-badge">✨ AI-POWERED</div>
        <h1 class="header-title">Resume Analyzer</h1>
        <p class="header-subtitle">Upload your resume and get intelligent insights in seconds</p>
        <div style="display:flex;gap:0.8rem;margin-top:1rem;flex-wrap:wrap;">
            <span style="background:rgba(255,255,255,0.08);padding:0.2rem 0.8rem;border-radius:100px;color:rgba(255,255,255,0.7);font-size:0.7rem;">📄 PDF</span>
            <span style="background:rgba(255,255,255,0.08);padding:0.2rem 0.8rem;border-radius:100px;color:rgba(255,255,255,0.7);font-size:0.7rem;">📝 DOCX</span>
            <span style="background:rgba(255,255,255,0.08);padding:0.2rem 0.8rem;border-radius:100px;color:rgba(255,255,255,0.7);font-size:0.7rem;">🤖 AI Analysis</span>
        </div>
    </div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("""
        <div style="text-align:center;padding:0.5rem 0;">
            <div style="font-size:2rem;">⚙️</div>
            <h4 style="margin:0;color:#1a1a2e;">Settings</h4>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    api_key = st.text_input(
        "🔑 Groq API Key",
        type="password",
        value="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"
    )
    
    if api_key:
        os.environ["GROQ_API_KEY"] = api_key
        st.success("✅ Connected")
    
    st.markdown("---")
    
    model_choice = st.selectbox(
        "🤖 Model",
        [
            "llama-3.3-70b-versatile",
            "llama-3.1-8b-instant",
            "meta-llama/llama-4-scout-17b-16e-instruct"
        ],
        index=0
    )
    
    st.markdown("---")
    
    st.markdown("""
        <div style="background:#f0f4ff;padding:1rem;border-radius:10px;">
            <p style="font-weight:600;color:#1a1a2e;margin:0 0 0.3rem 0;">💡 Tips</p>
            <ul style="color:#555;font-size:0.8rem;padding-left:1rem;margin:0;">
                <li>PDF/DOCX supported</li>
                <li>Add job description</li>
                <li>Get AI analysis</li>
            </ul>
        </div>
    """, unsafe_allow_html=True)

# Main Content
col1, col2 = st.columns([1, 1.2], gap="large")

with col1:
    st.markdown("""
        <div class="premium-card fade-in">
            <div class="card-title">📤 Upload</div>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader(
        " ",
        type=['pdf', 'docx'],
        label_visibility="collapsed"
    )
    
    if not uploaded_file:
        st.markdown("""
            <div class="upload-area">
                <div class="upload-icon">📄</div>
                <p style="font-weight:500;color:#1a1a2e;">Drop your resume here</p>
                <p style="color:#888;font-size:0.8rem;">PDF or DOCX</p>
            </div>
        """, unsafe_allow_html=True)
    
    job_description = st.text_area(
        "📋 Job Description",
        placeholder="Paste job description for better analysis...",
        height=100
    )
    
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown("""
        <div class="premium-card fade-in">
            <div class="card-title">📊 Results</div>
    """, unsafe_allow_html=True)
    
    if uploaded_file:
        st.markdown(f"""
            <div style="background:#f0f4ff;padding:0.6rem 1rem;border-radius:8px;margin-bottom:1rem;">
                <span style="font-weight:500;">📎 {uploaded_file.name}</span>
                <span style="color:#888;font-size:0.8rem;"> ({(uploaded_file.size/1024):.1f} KB)</span>
            </div>
        """, unsafe_allow_html=True)
        
        text = ""
        try:
            if uploaded_file.name.endswith('.pdf'):
                pdf_reader = PyPDF2.PdfReader(uploaded_file)
                for page in pdf_reader.pages:
                    text += page.extract_text()
            else:
                doc = docx.Document(uploaded_file)
                for para in doc.paragraphs:
                    text += para.text + "\n"
            
            # Extract skills
            skills_list = [
                'python', 'java', 'javascript', 'react', 'angular', 'vue',
                'aws', 'azure', 'gcp', 'sql', 'mongodb', 'postgresql', 'mysql',
                'html', 'css', 'docker', 'kubernetes', 'git', 'linux',
                'node.js', 'django', 'flask', 'spring', 'c++', 'c#',
                'machine learning', 'ai', 'nlp', 'data science'
            ]
            
            found_skills = []
            for skill in skills_list:
                if skill in text.lower():
                    found_skills.append(skill)
            
            if found_skills:
                st.markdown("**🔹 Skills**")
                skills_html = " ".join([f'<span class="skill-tag">{s}</span>' for s in found_skills[:12]])
                st.markdown(f'<div style="margin:0.5rem 0 1rem 0;">{skills_html}</div>', unsafe_allow_html=True)
            
            # Education
            edu_keywords = ['b.tech', 'bachelor', 'master', 'degree', 'university', 'college']
            sentences = text.split('.')
            edu_found = []
            for sentence in sentences:
                for keyword in edu_keywords:
                    if keyword in sentence.lower():
                        edu_found.append(sentence.strip())
                        break
            
            if edu_found:
                st.markdown("**🎓 Education**")
                for edu in edu_found[:2]:
                    st.markdown(f'<div class="info-item">{edu}</div>', unsafe_allow_html=True)
            
            if api_key:
                if st.button("✨ Generate AI Analysis", use_container_width=True):
                    with st.spinner("🧠 Analyzing..."):
                        try:
                            client = Groq(api_key=api_key)
                            
                            prompt = f"""
                            Analyze this resume and provide:
                            
                            1. Profile Summary (2-3 sentences)
                            2. Key Strengths (3-4 bullet points)
                            3. Areas for Improvement (2-3 bullet points)
                            4. Suggested Job Roles (3-4 roles)
                            
                            Resume:
                            {text[:3000]}
                            """
                            
                            if job_description:
                                prompt += f"\n\nJob Description:\n{job_description}"
                            
                            response = client.chat.completions.create(
                                model=model_choice,
                                messages=[
                                    {"role": "system", "content": "You are a professional resume analyst. Provide clear, actionable feedback."},
                                    {"role": "user", "content": prompt}
                                ],
                                temperature=0.7,
                                max_tokens=600
                            )
                            
                            st.markdown("---")
                            st.markdown("### 🤖 AI Analysis")
                            
                            # Display with proper styling for readability
                            analysis_text = response.choices[0].message.content
                            st.markdown(f'<div class="analysis-result">{analysis_text}</div>', unsafe_allow_html=True)
                            
                        except Exception as e:
                            st.error(f"Error: {str(e)[:100]}")
                            st.info("Try a different model from the sidebar")
            else:
                st.warning("⚠️ Add Groq API key in sidebar")
                
        except Exception as e:
            st.error(f"Error: {str(e)[:100]}")
    else:
        st.markdown("""
            <div style="text-align:center;padding:3rem 1rem;">
                <div style="font-size:4rem;">📄</div>
                <h4 style="color:#1a1a2e;">Upload a resume</h4>
                <p style="color:#888;font-size:0.85rem;">
                    PDF or DOCX · Free AI analysis
                </p>
            </div>
        """, unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown("""
    <div style="text-align:center;padding:1.5rem 0;color:#888;font-size:0.8rem;">
        🚀 Powered by Groq · Made with ❤️
    </div>
""", unsafe_allow_html=True)