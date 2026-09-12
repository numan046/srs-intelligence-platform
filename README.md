# SRS Intelligence Platform

<div align="center">

**AI-Powered Software Requirements Specification Analysis & Project Planning**

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.36+-red.svg)](https://streamlit.io)
[![OpenAI](https://img.shields.io/badge/OpenAI-GPT--4o--mini-green.svg)](https://openai.com)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

</div>

---

## Overview

SRS Intelligence Platform is an enterprise-grade web application that automatically analyzes Software Requirements Specification (SRS) documents using Artificial Intelligence. It transforms weeks of manual requirements review into minutes of intelligent automation, providing comprehensive analysis, quality validation, and actionable project planning insights.

## Features

### Core Capabilities

- **Smart Document Processing**
  - Support for PDF, DOCX, and TXT formats
  - Automatic text extraction and cleaning
  - Intelligent document chunking for large files

- **AI-Powered Analysis (18 Categories)**
  - Business Goals & Objectives
  - Functional & Non-Functional Requirements
  - Stakeholders & Constraints
  - Assumptions & Dependencies
  - Risks & Mitigation Strategies
  - Scope Definition
  - Missing Requirements Detection
  - Quality Issues Identification
  - Priority Classification
  - Timeline Estimation
  - Team Recommendations
  - Project Complexity Assessment

- **Requirement Quality Validation**
  - Automatic detection of ambiguous requirements
  - Incomplete requirement identification
  - Unmeasurable criteria flagging
  - Testability assessment
  - Conflict detection

- **Interactive AI Chat (RAG)**
  - Document-specific question answering
  - FAISS vector store for semantic search
  - Context-aware responses
  - Multi-turn conversation support

### Project Planning Modules

- **Developer Roadmap** - Phased development timeline
- **Cost Estimation** - Detailed budget breakdown
- **Tech Stack Advisor** - Technology recommendations
- **System Architect** - Architecture suggestions
- **Output Predictor** - Expected deliverables
- **Coverage Score** - Requirements completeness metric
- **SRS Improver** - Automated improvement suggestions
- **Use Case Generator** - Comprehensive use cases
- **Risk & Feasibility** - Risk assessment matrix
- **Team Roles** - Role & responsibility allocation
- **Project Generator** - Complete project structure
- **Management Dashboard** - Project overview
- **Export Center** - PDF reports & data export

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Frontend | Streamlit (Python) |
| Backend | Python 3.10+ |
| AI Engine | OpenAI GPT-4o-mini |
| Embeddings | OpenAI text-embedding-3-small |
| Vector Store | FAISS |
| Document Parsing | PyMuPDF, python-docx |
| Authentication | bcrypt |
| Report Generation | ReportLab |
| LangChain | LangChain + langchain-openai |

## Installation

### Prerequisites

- Python 3.10 or higher
- pip package manager
- OpenAI API key

### Setup Steps

1. **Clone the repository**
   ```bash
   git clone https://github.com/numan046/srs-intelligence-platform.git
   cd srs-intelligence-platform
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure API key**
   
   Create `.streamlit/secrets.toml`:
   ```toml
   OPENAI_API_KEY = "sk-your-api-key-here"
   ```

4. **Run the application**
   ```bash
   streamlit run app.py
   ```

5. **Access the app**
   
   Open your browser and navigate to `http://localhost:8501`

## Usage Guide

### 1. Authentication
- Sign up with your email and password
- Log in to access the dashboard

### 2. Upload SRS Document
- Navigate to "Upload SRS" in the sidebar
- Drag & drop or browse for your file (PDF/DOCX/TXT)
- Click "Analyze Document"
- Wait for AI analysis to complete (1-3 minutes)

### 3. View Analysis Results
- Dashboard shows comprehensive 18-category analysis
- Quality validation scores
- Improvement suggestions

### 4. Explore Planning Modules
- Access any module from the sidebar
- Each module provides AI-generated insights
- Export results as needed

### 5. Chat with Document
- Go to "AI Chat" section
- Ask questions about your SRS document
- Get instant, context-aware answers

## Project Structure

```
srs-intelligence-platform/
├── app.py                 # Main application (single-file Streamlit app)
├── requirements.txt       # Python dependencies
├── .gitignore            # Git ignore rules
├── .streamlit/
│   └── secrets.toml      # API keys (not committed to git)
└── data/                 # User data directory (auto-created)
    ├── users.json        # User accounts
    ├── history/          # Analysis history
    ├── vectors/          # FAISS vector stores
    ├── uploads/          # Uploaded documents
    └── reports/          # Generated reports
```

## Configuration

### Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `OPENAI_API_KEY` | OpenAI API key for GPT-4 and embeddings | Yes |

### Streamlit Secrets

Store sensitive configuration in `.streamlit/secrets.toml`:

```toml
OPENAI_API_KEY = "sk-..."
```

## Deployment

### Streamlit Cloud (Recommended)

1. Push code to GitHub
2. Visit [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repository
4. Add `OPENAI_API_KEY` in Secrets
5. Deploy!

Your app will be live at: `https://your-app-name.streamlit.app`

### Local Development

```bash
streamlit run app.py --server.port 8501
```

## Security Notes

- **Never commit** `.streamlit/secrets.toml` to version control
- API keys are stored securely in Streamlit secrets
- User passwords are hashed with bcrypt
- Each user has isolated data storage
- File uploads are stored in private user directories

## Troubleshooting

### Common Issues

**ModuleNotFoundError: No module named 'langchain.text_splitter'**
- Solution: Updated to `langchain_text_splitters` in newer versions

**OpenAI API key not configured**
- Solution: Add your key to `.streamlit/secrets.toml`

**Request Entity Too Large**
- Solution: Large documents are automatically truncated to 18,000 characters

**Port already in use**
- Solution: Use a different port: `streamlit run app.py --server.port 8502`

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Live Working App
https://srs-intelligence-platform-1234.streamlit.app/

## Author

**Muhammad Numan**
- GitHub: [@numan046](https://github.com/numan046)
- Email: numandine@gmail.com

## Acknowledgments

- OpenAI for GPT-4o-mini and embedding models
- Streamlit for the amazing web framework
- LangChain for LLM integration tools
- FAISS for efficient similarity search

---

<div align="center">

**Built with ❤️ for the software engineering community**

If this project helps you, please give it a ⭐ on GitHub!

</div>
