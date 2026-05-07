Custom MI Workflow & Task Management Application
"I build things that eliminate work that shouldn't exist."

This repository contains the front-end and back-end code for a bespoke, full-stack workflow management web application built entirely within the Google Workspace ecosystem.

Note: This is a sanitized version of a production application. All proprietary company data, internal URLs, and identifiable user information have been redacted or replaced with dummy variables.

📸 Application Interface
Markdown format: ![Main Workflow Dashboard](https://github.com/CraigPitt/team-workflow-management/blob/main/Workflowapp.png)

🚨 The Problem
For years, operational managers and analysts were losing hours of their week to a decentralized, highly manual reporting process. Work allocation was tracked across multiple disconnected spreadsheets, task handover (cover for absence) required manual tracking, and reminding users of their outstanding tasks involved drafting repetitive daily emails.

The business needed a centralized workflow tool, but off-the-shelf SaaS solutions were either too expensive or lacked the specific flexibility required for our internal processes.

💡 The Solution
I engineered a custom web application from the ground up using Google Apps Script, JavaScript (ES6+), HTML, and Tailwind CSS.

By using Google Sheets purely as a backend database and serving a custom HTML frontend, I created a zero-cost, highly polished application that completely replaced the manual spreadsheet process. What used to take hours of manual data entry and email chasing now happens automatically or via a single click.

🚀 Key Features
Real-Time Task Allocation: A dynamic, interactive workflow table where admins can assign reports and tasks to specific users.

Role-Based Access Control (RBAC): Secure login flow. "Admin" users see global management tools, while standard "Users" only see the tasks allocated specifically to them.

Automated Email Engine: * Instantly emails a user when a new task is assigned to them.

Sends automated reminders for overdue tasks.

Compiles and sends a daily summary email to leadership detailing outstanding vs. completed tasks.

Live KPI Dashboard: Integrates with Google Charts to provide real-time visual analytics on daily completion rates, team workload distribution, and on-time delivery (OTD) metrics.

Leave/Cover Management: A scheduling engine that allows admins to set temporary "Override Users" for tasks when colleagues are on annual leave.

Optimistic Data Locking: Prevents data collision by checking if a row has been modified by another user before executing a save.

📸 Live KPI Dashboard
Markdown format: ![KPI Dashboard](path/to/your/kpi-image.png)

💻 Tech Stack & Architecture
This application operates on a serverless architecture within Google Workspace:

Backend Database: Google Sheets (Holds user data, logs, and workflow status).

Backend Logic (ServerLogic.gs): Google Apps Script (V8 Runtime). Handles all API requests, data sorting, email dispatching, and security validation.

Frontend UI (index.html): HTML5, styled via Tailwind CSS via CDN for a modern, responsive interface.

Frontend Logic: Vanilla JavaScript (ES6+) handling DOM manipulation, asynchronous data fetching via google.script.run, and rendering Google Charts.

📂 Repository Structure
index.html: The complete frontend user interface and client-side JavaScript.

Server Logic.gs: The backend Apps Script containing the core CRUD operations, email triggers, and authorization logic.

🤝 Let's Connect
I am a self-taught Data Analyst specializing in SQL, Google BigQuery, and building custom automated infrastructure. I specialize in owning the full analytics lifecycle—from surfacing insights to building the tools that deliver them automatically.

LinkedIn: https://www.linkedin.com/in/craig-pitt/

Email: CraigPitt@mail.com
