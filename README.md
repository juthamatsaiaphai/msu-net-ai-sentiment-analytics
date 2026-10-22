# MSU Net AI Sentiment Analytics

**AI-Powered Sentiment Analysis and Workflow Automation for MSU Net / Wi-Fi at Mahasarakham University**

## 📌 About the Project

This research project focuses on analyzing user opinions, technology acceptance, and satisfaction with MSU Net / Wi-Fi services at Mahasarakham University.

The system uses Artificial Intelligence (AI) to analyze user feedback, identify key issues, summarize opinions, and generate suggestions for service improvement. Workflow automation with **n8n** helps connect data collection, AI-powered analysis, and reporting to support data-driven decision-making.

## 🎯 Objectives

- Analyze user opinions and satisfaction with MSU Net / Wi-Fi services.
- Classify user feedback into sentiment categories: Positive, Neutral, and Negative.
- Identify common connectivity issues and service improvement opportunities.
- Apply AI to summarize feedback and generate actionable suggestions.
- Use workflow automation to streamline data processing and reporting.
- Present analytical results to support data-driven decision-making.

## ⚙️ Technologies Used

- **Google Forms** — Collect survey responses.
- **Google Sheets** — Store and organize survey data.
- **n8n** — Automate the data processing workflow.
- **AI / Large Language Model (LLM)** — Analyze sentiment, summarize feedback, and generate suggestions.
- **Python / Streamlit** — Dashboard development, if used in the implemented system.
- **Data Visualization** — Present analytical findings and key insights.

## 🔄 System Workflow

1. Collect user feedback through Google Forms.
2. Store survey responses in Google Sheets.
3. Use n8n to automate the data processing workflow.
4. Send feedback to an AI service for sentiment analysis and summarization.
5. Organize the analysis results, key issues, and suggestions.
6. Present the results through a dashboard or reporting interface.

## 📊 Data Analysis

The project analyzes the following information:

- **User Profile:** Student or staff status and academic information.
- **Usage Behavior:** Frequency, device, location, duration, and time of use.
- **Technology Acceptance:** Perceived Usefulness (PU), Perceived Ease of Use (PEOU), Attitude Toward Using (ATU), and Behavioral Intention (BI).
- **Sentiment Analysis:** Positive, Neutral, and Negative feedback categories.
- **Key Issues:** Connection instability, Wi-Fi coverage, internet speed, and reconnection difficulties.
- **AI-Generated Insights:** Feedback summaries and suggestions for service improvement.

## 📁 Project Structure

```text
msu-net-ai-sentiment-analytics/
├── README.md
├── app.py                  # Dashboard application, if applicable
├── requirements.txt        # Python dependencies, if applicable
├── .gitignore
└── ...
```

The project structure may change as development progresses.

## 🚀 Getting Started

The setup instructions depend on the final project implementation.

For a Python-based dashboard, a typical setup is:

```bash
git clone https://github.com/juthamatsaiaphai/msu-net-ai-sentiment-analytics.git
cd msu-net-ai-sentiment-analytics
```

Create and activate a virtual environment, then install the dependencies if a `requirements.txt` file is available:

```bash
pip install -r requirements.txt
```

If the dashboard is implemented with Streamlit, run:

```bash
streamlit run app.py
```

Configure any required API credentials using environment variables or a local `.env` file. Do not commit API keys, passwords, or sensitive survey data to GitHub.

## 🔐 Data Privacy

Survey data may contain sensitive or personally identifiable information. Only publish anonymized data and ensure that any shared dataset is appropriate for public release.

API keys, passwords, and private credentials must not be included in the repository.

## 📌 Project Status

Research and system development project. Features and implementation details may be updated as development progresses.

## 👩‍💻 Author

**Juthamat Saiaphai**

Digital Business and Information Systems  
Faculty of Accountancy and Management  
Mahasarakham University

---

*This project explores the application of AI-powered analytics and workflow automation to understand user feedback and support service improvement.*
