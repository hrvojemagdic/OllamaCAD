
# 🚀 OllamaCAD
## AI-Powered SOLIDWORKS Add-in using Local LLMs & Vision Models

OllamaCAD is a production-ready SOLIDWORKS add-in that integrates local large language models (LLMs) and vision-language models (VLMs) directly inside the CAD environment.

The system enables context-aware AI assistance for mechanical design workflows without any cloud dependency.

## 🔥 Key Features
### 🧠 Local AI Integration (Ollama)

Fully offline inference

Supports LLM and VLM models (e.g. Qwen-VL, Gemma, etc.)

No data leakage

Compatible with NVIDIA GPUs for accelerated inference

### 🖼 Screenshot-Aware Assistance

Capture full SOLIDWORKS window

Send model view directly to vision models

Drawing grammar correction

Model inspection assistance

### 📦 Project-Based Persistent Memory

Per-CAD-file memory folder

Conversation history storage (JSONL)

Periodic summary compression via LLM

Screenshot archiving

Fully isolated project intelligence

### 📊 Assembly Intelligence Export

OpenXML-based Excel report generation

Mass properties, materials, bounding box

Feature statistics

Custom property editing via Excel

Import changes back into SOLIDWORKS

### 📚 RAG (Retrieval-Augmented Generation)

Per-project document ingestion

FAISS vector indexing

OCR via vision models

QA model for contextual answering

Fully local Python environment bootstrap

### ⚙️ Structured JSON Action Routing

LLM returns structured JSON

ActionRouter executes SOLIDWORKS commands

Foundation for autonomous CAD operations

## 🏗 Architecture Overview

OllamaCAD follows a modular architecture:

UI Layer → ChatPaneControl

LLM Layer → OllamaClient

Memory Layer → ProjectMemoryStore + Summarizer

RAG Engine → Python + FAISS

Execution Layer → ActionRouter

CAD Context Injection → SwSelectionProperties

This separation ensures extensibility and production-grade maintainability.

## 🖥 Requirements

SOLIDWORKS 2020+

.NET Framework 4.7.2+

Ollama running locally (http://localhost:11434)

Python 3.9+ (for RAG setup)

Optional: NVIDIA GPU for accelerated inference

## 🚀 Installation

Build project in Release | x64

Run OllamaCAD_Addin.reg

Start SOLIDWORKS

Enable "Ollama CAD" add-in

Ensure Ollama service is running

For RAG mode:

Click "Setup Global RAG Environment"

Place documents inside project OllamaRAG folder

Click "Build / Refresh RAG index"

## 🔐 Privacy & Security

Fully local processing

No external APIs

No telemetry

No cloud storage

Designed for enterprise-sensitive engineering environments.

## 🏆 Competition Context

This project was developed as an advanced engineering prototype for:

Ollama ecosystem integration

NVIDIA GPU acceleration potential

AI-native CAD workflow transformation

The focus is on:

Real-world mechanical design workflows

Offline AI deployment

Structured CAD automation

Practical engineering integration

## 📌 Roadmap

Expanded CAD action automation

Constraint-aware design modifications

Simulation-aware reasoning

Multi-document RAG graphs

Agentic workflow orchestration

## 📄 License
