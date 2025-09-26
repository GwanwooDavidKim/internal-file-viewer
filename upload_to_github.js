import { Octokit } from '@octokit/rest';
import fs from 'fs';
import path from 'path';

let connectionSettings;

async function getAccessToken() {
  if (connectionSettings && connectionSettings.settings.expires_at && new Date(connectionSettings.settings.expires_at).getTime() > Date.now()) {
    return connectionSettings.settings.access_token;
  }
  
  const hostname = process.env.REPLIT_CONNECTORS_HOSTNAME;
  const xReplitToken = process.env.REPL_IDENTITY 
    ? 'repl ' + process.env.REPL_IDENTITY 
    : process.env.WEB_REPL_RENEWAL 
    ? 'depl ' + process.env.WEB_REPL_RENEWAL 
    : null;

  if (!xReplitToken) {
    throw new Error('X_REPLIT_TOKEN not found for repl/depl');
  }

  connectionSettings = await fetch(
    'https://' + hostname + '/api/v2/connection?include_secrets=true&connector_names=github',
    {
      headers: {
        'Accept': 'application/json',
        'X_REPLIT_TOKEN': xReplitToken
      }
    }
  ).then(res => res.json()).then(data => data.items?.[0]);

  const accessToken = connectionSettings?.settings?.access_token || connectionSettings.settings?.oauth?.credentials?.access_token;

  if (!connectionSettings || !accessToken) {
    throw new Error('GitHub not connected');
  }
  return accessToken;
}

async function getUncachableGitHubClient() {
  const accessToken = await getAccessToken();
  return new Octokit({ auth: accessToken });
}

// 파일을 Base64로 인코딩하는 함수
function encodeFile(filePath) {
  const content = fs.readFileSync(filePath);
  return Buffer.from(content).toString('base64');
}

// 디렉토리의 모든 파일을 재귀적으로 가져오는 함수
function getAllFiles(dirPath, arrayOfFiles = []) {
  const files = fs.readdirSync(dirPath);

  files.forEach(file => {
    const fullPath = path.join(dirPath, file);
    if (fs.statSync(fullPath).isDirectory()) {
      // __pycache__ 등 제외
      if (!file.startsWith('.') && !file.includes('__pycache__')) {
        arrayOfFiles = getAllFiles(fullPath, arrayOfFiles);
      }
    } else {
      // 불필요한 파일 제외
      if (!file.startsWith('.') && 
          !file.endsWith('.pyc') && 
          file !== 'upload_to_github.js' &&
          file !== 'package.json' &&
          file !== 'package-lock.json') {
        arrayOfFiles.push(fullPath);
      }
    }
  });

  return arrayOfFiles;
}

async function createRepository() {
  try {
    const octokit = await getUncachableGitHubClient();
    
    // 리포지토리 정보  
    const repoName = 'korean-internal-file-viewer';
    const repoDescription = '🇰🇷 사내 파일 뷰어 - Korean Internal File Viewer: Desktop application for viewing and searching various business documents (PDF, PPT, Excel, Word, images) with PyQt6. Features COM-based PowerPoint processing and smart caching system.';
    
    console.log('🚀 GitHub 리포지토리 생성 중...');
    
    let repo;
    try {
      // 리포지토리 생성
      repo = await octokit.rest.repos.createForAuthenticatedUser({
        name: repoName,
        description: repoDescription,
        private: false, // public으로 설정
        has_issues: true,
        has_projects: true,
        has_wiki: true
      });
      
      console.log(`✅ 리포지토리 생성 완료: ${repo.data.html_url}`);
    } catch (createError) {
      if (createError.message.includes('already exists')) {
        console.log('💡 리포지토리가 이미 존재합니다. 기존 리포지토리를 사용합니다.');
        // 기존 리포지토리 정보 가져오기
        const { data: user } = await octokit.rest.users.getAuthenticated();
        repo = { 
          data: { 
            owner: { login: user.login }, 
            name: repoName,
            html_url: `https://github.com/${user.login}/${repoName}`
          } 
        };
      } else {
        throw createError;
      }
    }
    
    // 모든 파일 가져오기
    console.log('📁 파일 목록 수집 중...');
    const allFiles = getAllFiles('.');
    
    console.log(`📄 업로드할 파일 수: ${allFiles.length}개`);
    
    // 각 파일을 리포지토리에 업로드
    for (const filePath of allFiles) {
      try {
        const relativePath = filePath.replace('./', '');
        console.log(`⬆️  업로드 중: ${relativePath}`);
        
        // 바이너리 파일인지 확인
        const isBinary = ['.png', '.jpg', '.jpeg', '.gif', '.ico', '.pdf'].some(ext => 
          filePath.toLowerCase().endsWith(ext)
        );
        
        let content;
        if (isBinary) {
          content = encodeFile(filePath);
        } else {
          content = fs.readFileSync(filePath, 'utf8');
        }
        
        // 파일을 GitHub에 업로드
        await octokit.rest.repos.createOrUpdateFileContents({
          owner: repo.data.owner.login,
          repo: repoName,
          path: relativePath,
          message: `Add ${relativePath}`,
          content: isBinary ? content : Buffer.from(content, 'utf8').toString('base64')
        });
        
      } catch (error) {
        console.error(`❌ 파일 업로드 실패 (${filePath}):`, error.message);
      }
    }
    
    console.log('🎉 GitHub 업로드 완료!');
    console.log(`🔗 리포지토리 링크: ${repo.data.html_url}`);
    
  } catch (error) {
    console.error('❌ GitHub 업로드 실패:', error);
    
    if (error.message.includes('already exists')) {
      console.log('💡 리포지토리가 이미 존재합니다. 기존 리포지토리를 업데이트합니다.');
    }
  }
}

// 실행
createRepository();