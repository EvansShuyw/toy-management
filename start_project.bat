@echo off

REM 启动后端服务
cd backend
call venv\Scripts\activate
start "Backend" venv\Scripts\python -m uvicorn main:app --reload --host 0.0.0.0 --port 8000

REM 启动前端服务
cd ..\frontend
npm run dev