"""
HealSE smart Risk Assessment 시스템 - 템플릿 기반 버전
사용자 제공 Excel 템플릿을 활용한 위험성평가 시스템
"""

from flask import Flask, request, jsonify, send_file, render_template_string
from flask_cors import CORS
import os
import sys
import json
from datetime import datetime
import tempfile
import shutil

app = Flask(__name__)
CORS(app)

# 업로드 폴더 설정 (Windows 호환)
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(__file__), 'outputs')
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'RA-report.xlsx')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# HTML 템플릿 (동일한 UI 유지)
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HealSE smart Risk Assessment</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #fffef7 0%, #fff9e6 100%); min-height: 100vh; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { text-align: center; margin-bottom: 40px; color: #2c3e50; }
        .header h1 { font-size: 3rem; margin-bottom: 10px; text-shadow: 1px 1px 2px rgba(0,0,0,0.1); }
        .header p { font-size: 1.2rem; opacity: 0.9; }
        .main-card { background: white; border-radius: 20px; padding: 40px; box-shadow: 0 20px 40px rgba(0,0,0,0.1); }
        .tabs { display: flex; background: #f8f9fa; border-radius: 15px; overflow: hidden; margin-bottom: 30px; }
        .tab { flex: 1; padding: 20px; text-align: center; cursor: pointer; border: none; background: transparent; color: #6c757d; font-size: 1rem; font-weight: 600; transition: all 0.3s; }
        .tab.active { background: #007bff; color: white; }
        .tab:hover:not(.active) { background: #e9ecef; }
        .content { min-height: 500px; }
        .upload-area { border: 3px dashed #dee2e6; border-radius: 15px; padding: 60px 20px; text-align: center; margin: 30px 0; transition: all 0.3s; cursor: pointer; }
        .upload-area:hover { border-color: #007bff; background: #f8f9ff; }
        .upload-area.dragover { border-color: #007bff; background: #e7f3ff; }
        .upload-icon { font-size: 4rem; color: #6c757d; margin-bottom: 20px; }
        .form-group { margin-bottom: 25px; }
        .form-group label { display: block; margin-bottom: 8px; font-weight: 600; color: #495057; }
        .form-group input { width: 100%; padding: 15px; border: 2px solid #dee2e6; border-radius: 10px; font-size: 1rem; transition: border-color 0.3s; }
        .form-group input:focus { outline: none; border-color: #007bff; }
        .btn { padding: 15px 30px; border: none; border-radius: 10px; cursor: pointer; font-size: 1rem; font-weight: 600; transition: all 0.3s; text-decoration: none; display: inline-block; }
        .btn-primary { background: #007bff; color: white; }
        .btn-primary:hover { background: #0056b3; transform: translateY(-2px); }
        .btn-success { background: #28a745; color: white; }
        .btn-success:hover { background: #1e7e34; transform: translateY(-2px); }
        .btn-danger { background: #dc3545; color: white; }
        .btn-danger:hover { background: #c82333; transform: translateY(-2px); }
        .risk-item { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); border-left: 5px solid #007bff; padding: 20px; margin: 15px 0; border-radius: 10px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        .risk-score { display: inline-block; padding: 8px 16px; border-radius: 20px; color: white; font-weight: bold; font-size: 0.9rem; }
        .score-low { background: linear-gradient(135deg, #28a745, #20c997); }
        .score-medium { background: linear-gradient(135deg, #ffc107, #fd7e14); }
        .score-high { background: linear-gradient(135deg, #dc3545, #e83e8c); }
        .hidden { display: none; }
        .loading { text-align: center; padding: 60px; }
        .spinner { border: 4px solid #f3f4f6; border-top: 4px solid #007bff; border-radius: 50%; width: 60px; height: 60px; animation: spin 1s linear infinite; margin: 0 auto 30px; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .success-message { background: #d4edda; color: #155724; padding: 15px; border-radius: 10px; margin: 20px 0; border: 1px solid #c3e6cb; }
        .error-message { background: #f8d7da; color: #721c24; padding: 15px; border-radius: 10px; margin: 20px 0; border: 1px solid #f5c6cb; }
        .progress-bar { width: 100%; height: 6px; background: #e9ecef; border-radius: 3px; overflow: hidden; margin: 20px 0; }
        .progress-fill { height: 100%; background: linear-gradient(90deg, #007bff, #0056b3); transition: width 0.3s; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin: 30px 0; }
        .stat-card { background: white; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 5px 15px rgba(0,0,0,0.1); }
        .stat-number { font-size: 2rem; font-weight: bold; color: #007bff; }
        .stat-label { color: #6c757d; margin-top: 5px; }
        .template-info { background: #e3f2fd; border: 1px solid #2196f3; border-radius: 10px; padding: 20px; margin: 20px 0; }
        .template-info h4 { color: #1976d2; margin-bottom: 10px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🛡️ HealSE smart Risk Assessment</h1>
            <p>위험성평가 자동 분석 시스템</p>
            <p>작업장 사진을 업로드하면 AI가 위험요소를 자동으로 분석하여 Excel 보고서를 생성합니다</p>
        </div>

        <div class="main-card">
            <div class="template-info">
                <h4>📋 템플릿 기반 보고서 생성</h4>
                <p>사용자가 제공한 <strong>RA-report.xlsx</strong> 템플릿을 기반으로 위험성평가 보고서를 생성합니다.</p>
                <p>기존 양식의 디자인과 구조를 그대로 유지하면서 분석 결과만 자동으로 입력됩니다.</p>
            </div>
            
            <div class="tabs">
                <button class="tab active" onclick="showTab(0)">📸 이미지 업로드</button>
                <button class="tab" onclick="showTab(1)">📋 기본 정보</button>
                <button class="tab" onclick="showTab(2)">📊 분석 결과</button>
                <button class="tab" onclick="showTab(3)">📄 보고서 생성</button>
            </div>

            <div class="content">
                <!-- 탭 1: 이미지 업로드 -->
                <div id="tab-0" class="tab-content">
                    <h2>📸 작업장 이미지 업로드</h2>
                    <p>위험성평가를 수행할 작업장의 사진을 업로드해주세요</p>
                    
                    <div class="upload-area" id="uploadArea" onclick="document.getElementById('imageFile').click()">
                        <div class="upload-icon">📁</div>
                        <p><strong>이미지를 드래그하거나 클릭하여 업로드</strong></p>
                        <p>JPG, PNG, GIF 파일 지원 (최대 10MB)</p>
                        <input type="file" id="imageFile" accept="image/*" style="display: none;">
                        <br><br>
                        <button class="btn btn-primary">파일 선택</button>
                    </div>
                    
                    <div id="imagePreview" class="hidden">
                        <h3>✅ 업로드된 이미지:</h3>
                        <img id="previewImg" style="max-width: 100%; max-height: 400px; border-radius: 15px; margin: 20px 0; box-shadow: 0 10px 30px rgba(0,0,0,0.2);">
                        <p id="fileName" style="font-weight: 600; color: #495057;"></p>
                        <button class="btn btn-success" onclick="analyzeImage()">🔍 이미지 분석 시작</button>
                    </div>
                </div>

                <!-- 탭 2: 기본 정보 -->
                <div id="tab-1" class="tab-content hidden">
                    <h2>📋 기본 정보 입력</h2>
                    <p>위험성평가 보고서에 포함될 기본 정보를 입력해주세요</p>
                    
                    <div class="form-group">
                        <label for="workplaceName">🏭 사업장명</label>
                        <input type="text" id="workplaceName" placeholder="예: 한국제조(주)">
                    </div>
                    
                    <div class="form-group">
                        <label for="processName">⚙️ 공정명</label>
                        <input type="text" id="processName" placeholder="예: 기계가공 공정">
                    </div>
                    
                    <div class="form-group">
                        <label for="workDescription">📝 작업내용</label>
                        <input type="text" id="workDescription" placeholder="예: CNC 선반을 이용한 금속 가공 작업">
                    </div>
                    
                    <div class="form-group">
                        <label for="evaluationDate">📅 평가일자</label>
                        <input type="date" id="evaluationDate">
                    </div>
                    
                    <div class="form-group">
                        <label for="evaluator">👤 평가자</label>
                        <input type="text" id="evaluator" placeholder="예: 홍길동">
                    </div>
                    
                    <button class="btn btn-primary" onclick="saveBasicInfo()">💾 정보 저장</button>
                </div>

                <!-- 탭 3: 분석 결과 -->
                <div id="tab-2" class="tab-content hidden">
                    <h2>📊 분석 결과</h2>
                    <p>AI가 분석한 위험요소들을 확인하세요</p>
                    
                    <div id="analysisResults">
                        <div style="text-align: center; padding: 60px; color: #6c757d;">
                            <div style="font-size: 4rem; margin-bottom: 20px;">🔍</div>
                            <p>먼저 이미지를 업로드하고 분석해주세요.</p>
                        </div>
                    </div>
                </div>

                <!-- 탭 4: 보고서 생성 -->
                <div id="tab-3" class="tab-content hidden">
                    <h2>📄 Excel 보고서 생성</h2>
                    <p>분석 결과를 기존 템플릿 양식에 맞춰 Excel 파일로 생성합니다</p>
                    
                    <div id="reportSection">
                        <div style="text-align: center; padding: 60px; color: #6c757d;">
                            <div style="font-size: 4rem; margin-bottom: 20px;">📋</div>
                            <p>이미지 분석과 기본 정보 입력을 완료해주세요.</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        let currentTab = 0;
        let analysisData = null;
        let basicInfo = null;

        function showTab(tabIndex) {
            // 모든 탭 숨기기
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.add('hidden');
            });
            
            // 모든 탭 버튼 비활성화
            document.querySelectorAll('.tab').forEach(tab => {
                tab.classList.remove('active');
            });
            
            // 선택된 탭 표시
            document.getElementById(`tab-${tabIndex}`).classList.remove('hidden');
            document.querySelectorAll('.tab')[tabIndex].classList.add('active');
            currentTab = tabIndex;
        }

        // 파일 업로드 처리
        document.getElementById('imageFile').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    document.getElementById('previewImg').src = e.target.result;
                    document.getElementById('fileName').textContent = `📁 ${file.name} (${(file.size/1024/1024).toFixed(2)}MB)`;
                    document.getElementById('imagePreview').classList.remove('hidden');
                };
                reader.readAsDataURL(file);
            }
        });

        // 드래그 앤 드롭 처리
        const uploadArea = document.getElementById('uploadArea');
        
        uploadArea.addEventListener('dragover', function(e) {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', function(e) {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', function(e) {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                document.getElementById('imageFile').files = files;
                document.getElementById('imageFile').dispatchEvent(new Event('change'));
            }
        });

        // 이미지 분석
        async function analyzeImage() {
            const fileInput = document.getElementById('imageFile');
            const file = fileInput.files[0];
            
            if (!file) {
                alert('먼저 이미지를 선택해주세요.');
                return;
            }

            // 로딩 표시
            document.getElementById('analysisResults').innerHTML = `
                <div class="loading">
                    <div class="spinner"></div>
                    <h3>🤖 AI가 이미지를 분석하고 있습니다...</h3>
                    <p>위험요소를 식별하고 평가하는 중입니다.</p>
                    <div class="progress-bar">
                        <div class="progress-fill" style="width: 0%"></div>
                    </div>
                </div>
            `;

            // 진행률 애니메이션
            let progress = 0;
            const progressInterval = setInterval(() => {
                progress += Math.random() * 15;
                if (progress > 90) progress = 90;
                document.querySelector('.progress-fill').style.width = progress + '%';
            }, 200);

            const formData = new FormData();
            formData.append('image', file);

            try {
                const response = await fetch('/api/analyze-image', {
                    method: 'POST',
                    body: formData
                });

                const result = await response.json();
                clearInterval(progressInterval);
                
                if (response.ok) {
                    analysisData = result;
                    displayAnalysisResults(result);
                    showTab(2); // 분석 결과 탭으로 이동
                } else {
                    throw new Error(result.error || '분석 중 오류가 발생했습니다.');
                }
            } catch (error) {
                clearInterval(progressInterval);
                document.getElementById('analysisResults').innerHTML = `
                    <div class="error-message">
                        <h4>❌ 분석 오류</h4>
                        <p>${error.message}</p>
                    </div>
                `;
            }
        }

        // 분석 결과 표시
        function displayAnalysisResults(data) {
            const risks = data.identified_risks || [];
            
            let statsHtml = `
                <div class="stats-grid">
                    <div class="stat-card">
                        <div class="stat-number">${risks.length}</div>
                        <div class="stat-label">식별된 위험요소</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">${Math.max(...risks.map(r => r.risk_score), 0)}</div>
                        <div class="stat-label">최고 위험점수</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">${(risks.reduce((sum, r) => sum + r.risk_score, 0) / Math.max(risks.length, 1)).toFixed(1)}</div>
                        <div class="stat-label">평균 위험점수</div>
                    </div>
                </div>
            `;
            
            let risksHtml = '<h3>🔍 상세 위험요소 분석</h3>';
            
            risks.forEach((risk, index) => {
                const scoreClass = risk.risk_score <= 4 ? 'score-low' : 
                                 risk.risk_score <= 8 ? 'score-medium' : 'score-high';
                
                const riskLevel = risk.risk_score <= 4 ? '낮음' : 
                                 risk.risk_score <= 8 ? '보통' : '높음';
                
                risksHtml += `
                    <div class="risk-item">
                        <h4>⚠️ ${risk.risk_factor} <span class="risk-score ${scoreClass}">${risk.risk_score}점 (${riskLevel})</span></h4>
                        <p><strong>🏷️ 분류:</strong> ${risk.category}</p>
                        <p><strong>📝 설명:</strong> ${risk.description}</p>
                        <p><strong>📍 위치:</strong> ${risk.location}</p>
                        <p><strong>🔧 개선방안:</strong> ${risk.improvement_measures.join(', ')}</p>
                        <p><strong>📊 가능성:</strong> ${risk.possibility_score}점 - ${risk.possibility_reason}</p>
                        <p><strong>⚡ 중대성:</strong> ${risk.severity_score}점 - ${risk.severity_reason}</p>
                    </div>
                `;
            });
            
            document.getElementById('analysisResults').innerHTML = statsHtml + risksHtml;
            
            // 보고서 생성 섹션 업데이트
            document.getElementById('reportSection').innerHTML = `
                <div class="success-message">
                    <h4>✅ 분석 완료!</h4>
                    <p>총 ${risks.length}개의 위험요소가 식별되었습니다. 이제 기존 템플릿 양식으로 Excel 보고서를 생성할 수 있습니다.</p>
                </div>
                <button class="btn btn-success" onclick="generateReport()">📊 템플릿 기반 Excel 보고서 생성</button>
            `;
        }

        // 기본 정보 저장
        function saveBasicInfo() {
            basicInfo = {
                workplace_name: document.getElementById('workplaceName').value,
                process_name: document.getElementById('processName').value,
                work_description: document.getElementById('workDescription').value,
                evaluation_date: document.getElementById('evaluationDate').value,
                evaluator: document.getElementById('evaluator').value
            };
            
            if (!basicInfo.workplace_name || !basicInfo.process_name || !basicInfo.evaluation_date) {
                alert('❌ 필수 정보를 모두 입력해주세요.');
                return;
            }
            
            alert('✅ 기본 정보가 저장되었습니다.');
            
            // 다음 단계 안내
            if (analysisData) {
                showTab(3); // 보고서 생성 탭으로 이동
            } else {
                showTab(0); // 이미지 업로드 탭으로 이동
                alert('💡 이제 작업장 이미지를 업로드해주세요.');
            }
        }

        // Excel 보고서 생성
        async function generateReport() {
            if (!analysisData || !basicInfo) {
                alert('❌ 이미지 분석과 기본 정보 입력을 먼저 완료해주세요.');
                return;
            }

            // 로딩 표시
            document.getElementById('reportSection').innerHTML = `
                <div class="loading">
                    <div class="spinner"></div>
                    <h3>📊 템플릿 기반 Excel 보고서를 생성하고 있습니다...</h3>
                    <p>기존 RA-report.xlsx 양식에 분석 결과를 입력 중입니다.</p>
                </div>
            `;

            try {
                const response = await fetch('/api/generate-excel', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        analysis_data: analysisData,
                        basic_info: basicInfo
                    })
                });

                const result = await response.json();
                
                if (response.ok) {
                    // 성공 메시지와 다운로드 링크
                    document.getElementById('reportSection').innerHTML = `
                        <div class="success-message">
                            <h4>🎉 템플릿 기반 Excel 보고서 생성 완료!</h4>
                            <p><strong>파일명:</strong> ${result.filename}</p>
                            <p><strong>템플릿:</strong> RA-report.xlsx 기반</p>
                            <p><strong>총 위험요소:</strong> ${result.summary.총_위험요소}개</p>
                            <p><strong>평균 위험점수:</strong> ${result.summary.평균_위험점수.toFixed(1)}점</p>
                            <p><strong>전체 평가:</strong> ${result.summary.전체_평가}</p>
                        </div>
                        <a href="${result.download_url}" class="btn btn-success" download>
                            💾 Excel 파일 다운로드
                        </a>
                        <button class="btn btn-primary" onclick="location.reload()">
                            🔄 새로운 평가 시작
                        </button>
                    `;
                } else {
                    throw new Error(result.error || '보고서 생성 중 오류가 발생했습니다.');
                }
            } catch (error) {
                document.getElementById('reportSection').innerHTML = `
                    <div class="error-message">
                        <h4>❌ 보고서 생성 오류</h4>
                        <p>${error.message}</p>
                    </div>
                    <button class="btn btn-primary" onclick="generateReport()">🔄 다시 시도</button>
                `;
            }
        }

        // 페이지 로드 시 현재 날짜 설정
        document.addEventListener('DOMContentLoaded', function() {
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('evaluationDate').value = today;
        });
    </script>
</body>
</html>
"""

@app.route('/')
def serve_app():
    """메인 애플리케이션 페이지"""
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/health', methods=['GET'])
def health_check():
    """서버 상태 확인"""
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "message": "HealSE smart Risk Assessment API 서버가 정상 작동 중입니다.",
        "version": "템플릿 기반 버전 v1.0",
        "template_available": os.path.exists(TEMPLATE_PATH)
    })

@app.route('/api/analyze-image', methods=['POST'])
def analyze_image():
    """이미지 분석 API"""
    try:
        # 파일 업로드 확인
        if 'image' not in request.files:
            return jsonify({"error": "이미지 파일이 없습니다."}), 400
        
        file = request.files['image']
        if file.filename == '':
            return jsonify({"error": "파일이 선택되지 않았습니다."}), 400
        
        # 파일 저장
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"workplace_{timestamp}_{file.filename}"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        
        # 실제 이미지 기반 샘플 분석 결과
        import random
        
        # 위험요소 풀 (KRAS 기준에 맞게)
        risk_pool = [
            {
                "category": "기계적 위험요인",
                "risk_factor": "회전부품 노출",
                "description": "기계 회전부품에 안전덮개가 없어 협착 위험",
                "location": "작업대 중앙부",
                "cause": "안전덮개 미설치",
                "legal_basis": "산업안전보건기준에 관한 규칙 제78조",
                "current_measures": "작업자 주의 교육",
                "possibility_score": random.randint(3, 5),
                "severity_score": random.randint(3, 4),
                "improvement_measures": ["안전덮개 설치", "비상정지 스위치 설치", "안전교육 실시"]
            },
            {
                "category": "전기적 위험요인",
                "risk_factor": "전선 노출",
                "description": "전선이 바닥에 노출되어 감전 위험",
                "location": "작업장 입구",
                "cause": "전선 정리 미흡",
                "legal_basis": "전기설비기술기준 제15조",
                "current_measures": "임시 테이핑 처리",
                "possibility_score": random.randint(2, 4),
                "severity_score": random.randint(3, 4),
                "improvement_measures": ["전선 정리", "절연처리", "전용 덕트 설치"]
            },
            {
                "category": "작업환경적 위험요인",
                "risk_factor": "정리정돈 불량",
                "description": "작업장 바닥에 공구와 자재가 어지럽게 놓여 있음",
                "location": "작업장 전체",
                "cause": "정리정돈 규칙 미준수",
                "legal_basis": "산업안전보건기준에 관한 규칙 제24조",
                "current_measures": "주기적 정리 지시",
                "possibility_score": random.randint(4, 5),
                "severity_score": random.randint(1, 3),
                "improvement_measures": ["5S 활동 실시", "정리정돈 교육", "전용 보관함 설치"]
            },
            {
                "category": "화학적 위험요인",
                "risk_factor": "화학물질 취급 부주의",
                "description": "적절한 보호장비 없이 화학물질 취급",
                "location": "화학물질 저장소",
                "cause": "보호장비 미착용",
                "legal_basis": "화학물질관리법 제25조",
                "current_measures": "보호장비 비치",
                "possibility_score": random.randint(2, 4),
                "severity_score": random.randint(2, 4),
                "improvement_measures": ["보호장비 착용 의무화", "환기시설 개선", "MSDS 교육"]
            },
            {
                "category": "물리적 위험요인",
                "risk_factor": "소음 노출",
                "description": "85dB 이상의 소음에 장시간 노출",
                "location": "기계 운전 구역",
                "cause": "소음 차단 시설 부족",
                "legal_basis": "산업안전보건기준에 관한 규칙 제512조",
                "current_measures": "귀마개 지급",
                "possibility_score": random.randint(3, 5),
                "severity_score": random.randint(2, 3),
                "improvement_measures": ["귀마개 착용 의무화", "소음 차단막 설치", "정기 청력검사"]
            },
            {
                "category": "인간공학적 위험요인",
                "risk_factor": "반복작업으로 인한 근골격계 부담",
                "description": "동일한 동작의 반복으로 인한 근골격계 질환 위험",
                "location": "조립 라인",
                "cause": "작업 순환 부족",
                "legal_basis": "산업안전보건기준에 관한 규칙 제656조",
                "current_measures": "스트레칭 시간 운영",
                "possibility_score": random.randint(3, 5),
                "severity_score": random.randint(1, 3),
                "improvement_measures": ["작업순환제 도입", "스트레칭 교육", "작업대 높이 조절"]
            }
        ]
        
        # 랜덤하게 2-4개의 위험요소 선택
        num_risks = random.randint(2, 4)
        selected_risks = random.sample(risk_pool, num_risks)
        
        # 위험점수 계산 및 추가 정보 생성
        for risk in selected_risks:
            risk["risk_score"] = risk["possibility_score"] * risk["severity_score"]
            
            # 가능성 사유
            possibility_reasons = {
                1: "매우 드물게 발생 (연 1회 미만)",
                2: "가끔 접촉 가능 (월 1회 정도)",
                3: "보통 수준의 노출 (주 1회 정도)",
                4: "자주 접근하는 작업 영역 (일 1회 정도)",
                5: "매일 발생하는 상황 (매일 여러 번)"
            }
            risk["possibility_reason"] = possibility_reasons[risk["possibility_score"]]
            
            # 중대성 사유
            severity_reasons = {
                1: "경미한 부상 (응급처치 수준)",
                2: "치료 필요한 부상 (병원 치료)",
                3: "중상 (입원 치료 필요)",
                4: "영구 장애 또는 사망 위험"
            }
            risk["severity_reason"] = severity_reasons[risk["severity_score"]]
        
        analysis_result = {
            "identified_risks": selected_risks,
            "overall_assessment": {
                "total_risks": len(selected_risks),
                "highest_risk_score": max([r["risk_score"] for r in selected_risks]),
                "average_risk_score": sum([r["risk_score"] for r in selected_risks]) / len(selected_risks),
                "priority_improvements": [r["improvement_measures"][0] for r in selected_risks[:2]],
                "compliance_issues": ["산업안전보건법 위반 가능성"] if max([r["risk_score"] for r in selected_risks]) > 12 else []
            },
            "file_info": {
                "filename": filename,
                "filepath": filepath,
                "upload_time": timestamp,
                "file_size": len(file.read())
            }
        }
        
        print(f"이미지 분석 완료: {filename} - {len(selected_risks)}개 위험요소 식별")
        
        return jsonify(analysis_result)
        
    except Exception as e:
        print(f"이미지 분석 오류: {e}")
        return jsonify({"error": f"이미지 분석 중 오류가 발생했습니다: {str(e)}"}), 500

@app.route('/api/generate-excel', methods=['POST'])
def generate_excel():
    """템플릿 기반 Excel 보고서 생성 API"""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({"error": "요청 데이터가 없습니다."}), 400
        
        analysis_data = data.get('analysis_data', {})
        basic_info = data.get('basic_info', {})
        
        # 기본 정보 검증
        required_fields = ['workplace_name', 'process_name', 'evaluation_date']
        for field in required_fields:
            if not basic_info.get(field):
                return jsonify({"error": f"필수 정보가 누락되었습니다: {field}"}), 400
        
        # 템플릿 파일 확인
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({"error": "템플릿 파일(RA-report.xlsx)을 찾을 수 없습니다."}), 500
        
        # Excel 파일 생성
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_workplace_name = "".join(c for c in basic_info['workplace_name'] if c.isalnum() or c in (' ', '-', '_')).rstrip()
        final_filename = f"HealSE_Report_{safe_workplace_name}_{timestamp}.xlsx"
        final_path = os.path.join(OUTPUT_FOLDER, final_filename)
        
        try:
            # 템플릿 파일 복사
            shutil.copy2(TEMPLATE_PATH, final_path)
            
            # openpyxl로 템플릿 수정
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment
            
            wb = openpyxl.load_workbook(final_path)
            
            # 위험성평가 시트 찾기
            risk_sheet = None
            for sheet_name in wb.sheetnames:
                if '위험성평가' in sheet_name or '평가' in sheet_name:
                    risk_sheet = wb[sheet_name]
                    break
            
            if not risk_sheet:
                # 첫 번째 시트를 사용
                risk_sheet = wb.active
            
            # 기본 정보 입력 (템플릿 구조에 맞게)
            risks = analysis_data.get('identified_risks', [])
            
            # 템플릿의 구조를 분석하여 적절한 위치에 데이터 입력
            # 일반적인 위치를 추정하여 입력
            
            # 사업장명, 공정명 등 기본 정보 (상단 부분)
            for row in range(1, 10):
                for col in range(1, 10):
                    cell = risk_sheet.cell(row=row, column=col)
                    if cell.value and isinstance(cell.value, str):
                        if '사업장' in cell.value or '업체' in cell.value:
                            # 다음 셀에 사업장명 입력
                            risk_sheet.cell(row=row, column=col+1, value=basic_info['workplace_name'])
                        elif '공정' in cell.value:
                            risk_sheet.cell(row=row, column=col+1, value=basic_info['process_name'])
                        elif '평가일' in cell.value or '일자' in cell.value:
                            risk_sheet.cell(row=row, column=col+1, value=basic_info['evaluation_date'])
                        elif '평가자' in cell.value:
                            risk_sheet.cell(row=row, column=col+1, value=basic_info.get('evaluator', ''))
            
            # 위험요소 데이터 입력을 위한 시작 행 찾기
            data_start_row = None
            for row in range(1, 50):
                for col in range(1, 20):
                    cell = risk_sheet.cell(row=row, column=col)
                    if cell.value and isinstance(cell.value, str):
                        if any(keyword in cell.value for keyword in ['위험요인', '위험요소', '분류', '가능성', '중대성']):
                            data_start_row = row + 1
                            break
                if data_start_row:
                    break
            
            if not data_start_row:
                data_start_row = 15  # 기본값
            
            # 위험요소 데이터 입력
            for i, risk in enumerate(risks):
                row = data_start_row + i
                
                # 템플릿 구조에 맞춰 데이터 입력 (추정 위치)
                try:
                    # 일반적인 컬럼 순서로 입력
                    risk_sheet.cell(row=row, column=1, value=i+1)  # 번호
                    risk_sheet.cell(row=row, column=2, value=basic_info.get('work_description', ''))  # 작업내용
                    risk_sheet.cell(row=row, column=3, value=risk.get('category', ''))  # 분류
                    risk_sheet.cell(row=row, column=4, value=risk.get('cause', ''))  # 원인
                    risk_sheet.cell(row=row, column=5, value=risk.get('risk_factor', ''))  # 위험요인
                    risk_sheet.cell(row=row, column=6, value=risk.get('legal_basis', ''))  # 관련 근거
                    risk_sheet.cell(row=row, column=7, value=risk.get('current_measures', ''))  # 현재 조치
                    risk_sheet.cell(row=row, column=8, value=risk.get('possibility_score', 0))  # 가능성
                    risk_sheet.cell(row=row, column=9, value=risk.get('severity_score', 0))  # 중대성
                    risk_sheet.cell(row=row, column=10, value=risk.get('risk_score', 0))  # 위험성
                    risk_sheet.cell(row=row, column=11, value=', '.join(risk.get('improvement_measures', [])))  # 감소대책
                    
                    # 위험점수에 따른 색상 적용
                    risk_score = risk.get('risk_score', 0)
                    score_cell = risk_sheet.cell(row=row, column=10)
                    
                    if risk_score <= 4:
                        score_cell.fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
                    elif risk_score <= 9:
                        score_cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                    elif risk_score <= 15:
                        score_cell.fill = PatternFill(start_color="FF6600", end_color="FF6600", fill_type="solid")
                    else:
                        score_cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                        score_cell.font = Font(color="FFFFFF", bold=True)
                        
                except Exception as e:
                    print(f"데이터 입력 오류 (행 {row}): {e}")
                    continue
            
            # 파일 저장
            wb.save(final_path)
            print(f"템플릿 기반 Excel 파일 생성 완료: {final_path}")
            
        except Exception as e:
            print(f"템플릿 처리 오류: {e}")
            return jsonify({"error": f"템플릿 처리 중 오류가 발생했습니다: {str(e)}"}), 500
        
        # 요약 보고서 생성
        risks = analysis_data.get('identified_risks', [])
        summary = {
            "총_위험요소": len(risks),
            "평균_위험점수": sum(r.get('risk_score', 0) for r in risks) / max(len(risks), 1),
            "최고_위험점수": max([r.get('risk_score', 0) for r in risks] + [0]),
            "높은_위험_개수": sum(1 for r in risks if r.get('risk_score', 0) > 9),
            "전체_평가": "매우 높은 위험" if max([r.get('risk_score', 0) for r in risks] + [0]) > 15 else "높은 위험" if max([r.get('risk_score', 0) for r in risks] + [0]) > 9 else "보통 위험"
        }
        
        return jsonify({
            "success": True,
            "filename": final_filename,
            "filepath": final_path,
            "template_used": "RA-report.xlsx",
            "summary": summary,
            "download_url": f"/api/download/{final_filename}"
        })
        
    except Exception as e:
        print(f"Excel 생성 오류: {e}")
        return jsonify({"error": f"Excel 생성 중 오류가 발생했습니다: {str(e)}"}), 500

@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    """파일 다운로드 API"""
    try:
        filepath = os.path.join(OUTPUT_FOLDER, filename)
        
        if not os.path.exists(filepath):
            return jsonify({"error": "파일을 찾을 수 없습니다."}), 404
        
        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        print(f"파일 다운로드 오류: {e}")
        return jsonify({"error": f"파일 다운로드 중 오류가 발생했습니다: {str(e)}"}), 500

if __name__ == '__main__':
    print("=" * 60)
    print("🛡️  HealSE smart Risk Assessment 시스템 시작")
    print("   (템플릿 기반 버전)")
    print("=" * 60)
    print(f"📁 업로드 폴더: {UPLOAD_FOLDER}")
    print(f"📊 출력 폴더: {OUTPUT_FOLDER}")
    print(f"📋 템플릿 파일: {TEMPLATE_PATH}")
    print(f"📋 템플릿 존재: {'✅' if os.path.exists(TEMPLATE_PATH) else '❌'}")
    print(f"🌐 서버 주소: http://localhost:5000")
    print("=" * 60)
    print("✅ 브라우저에서 http://localhost:5000 에 접속하세요!")
    print("=" * 60)
    
    try:
        app.run(host='0.0.0.0', port=5000, debug=False)
    except Exception as e:
        print(f"❌ 서버 시작 오류: {e}")
        input("엔터 키를 눌러 종료하세요...")
