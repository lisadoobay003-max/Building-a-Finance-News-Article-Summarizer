```python
!pip install -q newspaper3k python-dotenv lxml_html_clean langchain-openai python-docx
import json
from dotenv import load_dotenv
load_dotenv()
import requests
from newspaper import Article
!pip install langchain
from langchain_core.messages import HumanMessage
from langchain_openai import ChatOpenAI
import os
import textwrap
from docx import Document
from docx.shared import Inches
```

    [?25l   [90m━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━[0m [32m0.0/253.0 kB[0m [31m?[0m eta [36m-:--:--[0m[2K   [91m━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━[0m[91m╸[0m[90m━[0m [32m245.8/253.0 kB[0m [31m7.7 MB/s[0m eta [36m0:00:01[0m[2K   [90m━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━[0m [32m253.0/253.0 kB[0m [31m5.4 MB/s[0m eta [36m0:00:00[0m
    [?25hRequirement already satisfied: langchain in /usr/local/lib/python3.12/dist-packages (1.3.1)
    Requirement already satisfied: langchain-core<2.0.0,>=1.4.0 in /usr/local/lib/python3.12/dist-packages (from langchain) (1.4.0)
    Requirement already satisfied: langgraph<1.3.0,>=1.2.0 in /usr/local/lib/python3.12/dist-packages (from langchain) (1.2.1)
    Requirement already satisfied: pydantic<3.0.0,>=2.7.4 in /usr/local/lib/python3.12/dist-packages (from langchain) (2.12.3)
    Requirement already satisfied: jsonpatch<2.0.0,>=1.33.0 in /usr/local/lib/python3.12/dist-packages (from langchain-core<2.0.0,>=1.4.0->langchain) (1.33)
    Requirement already satisfied: langchain-protocol>=0.0.14 in /usr/local/lib/python3.12/dist-packages (from langchain-core<2.0.0,>=1.4.0->langchain) (0.0.15)
    Requirement already satisfied: langsmith<1.0.0,>=0.3.45 in /usr/local/lib/python3.12/dist-packages (from langchain-core<2.0.0,>=1.4.0->langchain) (0.8.5)
    Requirement already satisfied: packaging>=23.2.0 in /usr/local/lib/python3.12/dist-packages (from langchain-core<2.0.0,>=1.4.0->langchain) (26.2)
    Requirement already satisfied: pyyaml<7.0.0,>=5.3.0 in /usr/local/lib/python3.12/dist-packages (from langchain-core<2.0.0,>=1.4.0->langchain) (6.0.3)
    Requirement already satisfied: tenacity!=8.4.0,<10.0.0,>=8.1.0 in /usr/local/lib/python3.12/dist-packages (from langchain-core<2.0.0,>=1.4.0->langchain) (9.1.4)
    Requirement already satisfied: typing-extensions<5.0.0,>=4.7.0 in /usr/local/lib/python3.12/dist-packages (from langchain-core<2.0.0,>=1.4.0->langchain) (4.15.0)
    Requirement already satisfied: uuid-utils<1.0,>=0.12.0 in /usr/local/lib/python3.12/dist-packages (from langchain-core<2.0.0,>=1.4.0->langchain) (0.16.0)
    Requirement already satisfied: langgraph-checkpoint<5.0.0,>=4.1.0 in /usr/local/lib/python3.12/dist-packages (from langgraph<1.3.0,>=1.2.0->langchain) (4.1.0)
    Requirement already satisfied: langgraph-prebuilt<1.2.0,>=1.1.0 in /usr/local/lib/python3.12/dist-packages (from langgraph<1.3.0,>=1.2.0->langchain) (1.1.0)
    Requirement already satisfied: langgraph-sdk<0.4.0,>=0.3.0 in /usr/local/lib/python3.12/dist-packages (from langgraph<1.3.0,>=1.2.0->langchain) (0.3.14)
    Requirement already satisfied: xxhash>=3.5.0 in /usr/local/lib/python3.12/dist-packages (from langgraph<1.3.0,>=1.2.0->langchain) (3.7.0)
    Requirement already satisfied: annotated-types>=0.6.0 in /usr/local/lib/python3.12/dist-packages (from pydantic<3.0.0,>=2.7.4->langchain) (0.7.0)
    Requirement already satisfied: pydantic-core==2.41.4 in /usr/local/lib/python3.12/dist-packages (from pydantic<3.0.0,>=2.7.4->langchain) (2.41.4)
    Requirement already satisfied: typing-inspection>=0.4.2 in /usr/local/lib/python3.12/dist-packages (from pydantic<3.0.0,>=2.7.4->langchain) (0.4.2)
    Requirement already satisfied: jsonpointer>=1.9 in /usr/local/lib/python3.12/dist-packages (from jsonpatch<2.0.0,>=1.33.0->langchain-core<2.0.0,>=1.4.0->langchain) (3.1.1)
    Requirement already satisfied: ormsgpack>=1.12.0 in /usr/local/lib/python3.12/dist-packages (from langgraph-checkpoint<5.0.0,>=4.1.0->langgraph<1.3.0,>=1.2.0->langchain) (1.12.2)
    Requirement already satisfied: httpx>=0.25.2 in /usr/local/lib/python3.12/dist-packages (from langgraph-sdk<0.4.0,>=0.3.0->langgraph<1.3.0,>=1.2.0->langchain) (0.28.1)
    Requirement already satisfied: orjson>=3.11.5 in /usr/local/lib/python3.12/dist-packages (from langgraph-sdk<0.4.0,>=0.3.0->langgraph<1.3.0,>=1.2.0->langchain) (3.11.9)
    Requirement already satisfied: requests-toolbelt>=1.0.0 in /usr/local/lib/python3.12/dist-packages (from langsmith<1.0.0,>=0.3.45->langchain-core<2.0.0,>=1.4.0->langchain) (1.0.0)
    Requirement already satisfied: requests>=2.0.0 in /usr/local/lib/python3.12/dist-packages (from langsmith<1.0.0,>=0.3.45->langchain-core<2.0.0,>=1.4.0->langchain) (2.32.4)
    Requirement already satisfied: zstandard>=0.23.0 in /usr/local/lib/python3.12/dist-packages (from langsmith<1.0.0,>=0.3.45->langchain-core<2.0.0,>=1.4.0->langchain) (0.25.0)
    Requirement already satisfied: anyio in /usr/local/lib/python3.12/dist-packages (from httpx>=0.25.2->langgraph-sdk<0.4.0,>=0.3.0->langgraph<1.3.0,>=1.2.0->langchain) (4.13.0)
    Requirement already satisfied: certifi in /usr/local/lib/python3.12/dist-packages (from httpx>=0.25.2->langgraph-sdk<0.4.0,>=0.3.0->langgraph<1.3.0,>=1.2.0->langchain) (2026.5.20)
    Requirement already satisfied: httpcore==1.* in /usr/local/lib/python3.12/dist-packages (from httpx>=0.25.2->langgraph-sdk<0.4.0,>=0.3.0->langgraph<1.3.0,>=1.2.0->langchain) (1.0.9)
    Requirement already satisfied: idna in /usr/local/lib/python3.12/dist-packages (from httpx>=0.25.2->langgraph-sdk<0.4.0,>=0.3.0->langgraph<1.3.0,>=1.2.0->langchain) (3.15)
    Requirement already satisfied: h11>=0.16 in /usr/local/lib/python3.12/dist-packages (from httpcore==1.*->httpx>=0.25.2->langgraph-sdk<0.4.0,>=0.3.0->langgraph<1.3.0,>=1.2.0->langchain) (0.16.0)
    Requirement already satisfied: charset_normalizer<4,>=2 in /usr/local/lib/python3.12/dist-packages (from requests>=2.0.0->langsmith<1.0.0,>=0.3.45->langchain-core<2.0.0,>=1.4.0->langchain) (3.4.7)
    Requirement already satisfied: urllib3<3,>=1.21.1 in /usr/local/lib/python3.12/dist-packages (from requests>=2.0.0->langsmith<1.0.0,>=0.3.45->langchain-core<2.0.0,>=1.4.0->langchain) (2.5.0)
    


```python
headers = {
    'User-Agent': 'MicrosoftEdge (Windows NT 11.0; Win64; x64) AppleWebkit/537.36(KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36'
}
```


```python
article_url ="https://www.ey.com/en_gr/insights/financial-services/how-artificial-intelligence-is-reshaping-the-financial-services-industry"
```


```python
session = requests.Session()

try:
  response = session.get(article_url, headers= headers, timeout=10)

  if response.status_code == 200:
    article= Article(article_url)
    article.download()
    article.parse()

    print(f"Title: {article.title}")
    print(f"Text: {article.text}")
  else:
    print(f"Failed to fetch article at {article_url}")
except Exception as e:
  print(f"Error occured while fetching article at {article_url}: {e}")
```

    Title: How artificial intelligence is reshaping the financial services industry
    Text: Generative AI is driving a profound transformation in financial services, fostering innovation and streamlining operations.
    
    With its broad applications, artificial intelligence is enhancing customer service, boosting risk management and reshaping capital markets.
    
    Balancing the opportunities and challenges of AI, the banking sector is on a strategic journey toward an AI-enabled future.
    
    In the dynamic world of financial services, artificial intelligence (AI), particularly Generative AI (GenAI), has become the linchpin of transformative change, redefining the operational and strategic horizons of the banking sector. GenAI’s capacity for creating new, original content is not merely an incremental advancement but a change in basic assumptions that is propelling banking toward a future ripe with innovation and efficiency
    
    
    
    GenAI models such as GPT, with its transformer architecture, mark a quantum leap from the AI of yesteryear, which primarily focused on understanding and processing information. Today, these models are the architects of text, images, code and more, initiating an era of unparalleled innovation in banking. The strategic deployment of GenAI is much more than a trend; it is a comprehensive reimagining of operations, product development and risk management, allowing banks to deliver personalized services and novel solutions while streamlining mundane tasks.
    
    The evolution of AI in banking has been nothing short of revolutionary, moving from foundational concepts to the creation of sophisticated, innovative applications.
    
    This transformation is apparent in the broad spectrum of available AI applications, from automated knowledge management to investment research and bespoke banking services, each underscoring the remarkable advancements and potential of GenAI. Major banks, especially those in North America, have been pioneers in this journey, making substantial investments in AI to spearhead innovation, talent development and operational transparency. Their investment strategies encompass a wide range of applications, including enhancement of fraud detection mechanisms and customer service chatbots. Their focus is on acquiring critical hardware, such as NVIDIA chips for AI processes, and making strategic investments in human and technological resources. The aim of refining existing processes is driving this strategic shift, combined with an ambition to explore and capitalize on high-impact AI use cases, balance potential benefits against risks, and scale innovative prototypes into robust solutions.
    

**Summarizing with OpenAI ChatGPT 4.0 -mini and Langchain**


```python
#web scraping
article_title = article.title
article_text = article.text
```


```python
#template for prompt 1
template = """You are a very good assistant that summarizes online articles.

Here's the article that you are required to summarize.

================

Title: {article_title}

{article_text}

================

Write a summary of the previous article.
"""

prompt = template.format(article_title= article.title, article_text = article.text)

messages = [HumanMessage(content=prompt)]
```


```python
os.environ["OPENAI_API_KEY"] = "Enter OPEN_API_KEY"
chat = ChatOpenAI(model_name= "gpt-4o", temperature = 1)
```


```python
#generate summary
synopsis = chat.invoke(messages)
print(textwrap.fill(synopsis.content, width=80))
```

    The article discusses how artificial intelligence, particularly Generative AI
    (GenAI), is significantly transforming the financial services industry. GenAI's
    ability to create new, original content is not just an advancement but a
    fundamental shift in the banking sector's operations and strategies. This
    transformation includes enhancing customer service, improving risk management,
    and reshaping capital markets. GenAI models, such as GPT, mark significant
    progress by enabling novel applications in banking, including personalized
    services and streamlined operations. Major banks, particularly in North America,
    are heavily investing in AI to drive innovation, develop talent, and ensure
    operational transparency. These investments focus on applications like fraud
    detection and customer service chatbots, utilizing critical hardware like NVIDIA
    chips for AI processes. The overarching goal is to refine existing processes,
    evaluate the benefits and risks of AI, and develop innovative AI solutions on a
    large scale.
    


```python
#template for prompt 2
template2 = """You are an advanced AI assistant that creates bulleted lists when summarizing online articles

Here's the article that you are required to summarize.

==================

Title: {article_title}

{article_text}

=================

Now, provide a summary of the article with bulleted points.
"""

#format prompt
prompt2 =  template2.format(article_title= article.title, article_text =  article.text)

#generate summary
summary = chat.invoke([HumanMessage(content=prompt2)])
print(textwrap.fill(summary.content, width=80))
```

    - **Title and Focus**: The article discusses how artificial intelligence (AI),
    especially Generative AI (GenAI), is reshaping the financial services industry.
    - **Impact on Financial Services**: AI is driving significant innovation and
    streamlining operations within the industry. - **Applications**: AI enhances
    customer service, improves risk management, and transforms capital markets. -
    **Strategic Shift**: The banking sector is on a path to an AI-enabled future,
    weighing both opportunities and challenges. - **GenAI's Role**: GenAI is a
    catalyst for change, transforming the banking industry with innovative models
    like GPT that go beyond previous AI capabilities. - **Operational and Strategic
    Changes**: GenAI is reimagining operations, product development, and risk
    management to deliver personalized services and streamline tasks. -
    **Revolutionary Evolution**: AI in banking has evolved from basic concepts to
    sophisticated applications, including automated knowledge management and
    customized banking services. - **Industry Leadership**: Major North American
    banks are leading the AI transformation, investing heavily in AI for innovation,
    talent development, and transparency. - **Investment Strategies**: Banks focus
    on enhancing fraud detection, improving customer service through chatbots,
    acquiring critical hardware like NVIDIA chips, and investing in both human and
    technological resources. - **Objective**: The goal is to refine processes,
    explore high-impact AI use cases, balance benefits against risks, and scale
    prototypes into robust solutions.
    


```python
#template for prompt 3

template3="""You are an advanced AI assistant that has the ability to summarize online articles into bulleted lists in French.


Here's the article that you are required to summarize.

===========================
Title: {article_title}

{article_text}
=================

Now, provide a summarized version of the article into a bulleted list format, in French.
"""

#format prompt
prompt3 = template3.format(article_title= article.title, article_text = article.text)

#generate summary
summary2 = chat.invoke([HumanMessage(content=prompt3)])
print(textwrap.fill(summary2.content))
```

    - L'intelligence artificielle générative transforme profondément les
    services financiers, favorisant l'innovation et rationalisant les
    opérations. - L'IA améliore le service client, renforce la gestion des
    risques et redéfinit les marchés de capitaux. - Le secteur bancaire
    progresse stratégiquement vers un avenir où l'IA joue un rôle central,
    en équilibrant opportunités et défis. - Dans les services financiers
    dynamiques, l'IA, et en particulier l'IA générative, redéfinit les
    horizons opérationnels et stratégiques. - Les modèles d'IA générative
    comme GPT marquent une avancée significative par rapport aux anciennes
    technologies, initiant une ère d'innovation sans précédent. - La mise
    en œuvre stratégique de l'IA générative réinvente les opérations, le
    développement de produits et la gestion des risques, offrant des
    services personnalisés et solutions novatrices. - L'évolution de l'IA
    dans le secteur bancaire est révolutionnaire, passant de concepts de
    base à des applications sophistiquées et innovantes. - Les
    applications d'IA couvrent une large gamme, de la gestion automatisée
    des connaissances à la recherche d'investissements et aux services
    bancaires sur mesure. - Les grandes banques, en particulier en
    Amérique du Nord, ont été des pionnières, investissant massivement
    dans l'IA pour favoriser l'innovation, le développement de talents et
    la transparence opérationnelle. - Les stratégies d'investissement des
    banques incluent l'amélioration des mécanismes de détection de fraude
    et des chatbots pour le service client. - Les banques se concentrent
    sur l'acquisition de matériel essentiel, comme les puces NVIDIA, et
    les investissements stratégiques dans les ressources humaines et
    technologiques. - L'objectif est d'affiner les processus existants,
    d'explorer des cas d'utilisation de l'IA à fort impact tout en
    équilibrant les avantages potentiels et les risques, et de transformer
    des prototypes innovants en solutions robustes.
    

**Export to Word Document**


```python
document = Document()

document.add_heading(article_title, level=1)
document.add_heading('Original Article URL', level=2)
document.add_paragraph(article_url)

document.add_heading('Summary (Plain Text)', level=2)
document.add_paragraph(synopsis.content)

document.add_heading('Summary (Bulleted List)', level=2)

document.add_paragraph(summary.content)

document.add_heading('Summary (Bulleted List in French)', level=2)

document.add_paragraph(summary2.content)

# Save the document
doc_filename = 'summarized_article.docx'
document.save(doc_filename)
print(f"Summarized article saved to {doc_filename}")
```

    Summarized article saved to summarized_article.docx
    
