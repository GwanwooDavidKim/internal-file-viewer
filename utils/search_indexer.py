# -*- coding: utf-8 -*-
"""
검색 인덱서 모듈 (Search Indexer)

파일 내용을 인덱싱하고 빠른 전문 검색 기능을 제공합니다.
"""
import os
import re
import json
import time
import threading
import hashlib
from datetime import datetime
from typing import Dict, List, Any, Optional, Set, Tuple
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
import config
from utils.file_manager import FileManager


class SearchIndex:
    """
    검색 인덱스를 관리하는 클래스입니다.
    
    파일의 내용을 토큰화하여 역 인덱스(inverted index)를 구축하고,
    빠른 전문 검색을 지원합니다.
    """
    
    def __init__(self):
        """SearchIndex 인스턴스를 초기화합니다."""
        self.index = defaultdict(set)  # 단어 -> 파일 경로 집합
        self.file_info = {}  # 파일 경로 -> 파일 정보
        self.stop_words = self._load_stop_words()
        self.lock = threading.RLock()
    
    def _load_stop_words(self) -> Set[str]:
        """불용어 목록을 로드합니다."""
        # 한국어와 영어 기본 불용어
        korean_stop_words = {
            '이', '그', '저', '것', '의', '가', '을', '를', '에', '에서', '로', '으로',
            '은', '는', '이다', '있다', '하다', '되다', '수', '등', '및', '또는',
            '그리고', '하지만', '그러나', '따라서', '그래서'
        }
        
        english_stop_words = {
            'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'for', 'from',
            'has', 'he', 'in', 'is', 'it', 'its', 'of', 'on', 'that', 'the',
            'to', 'was', 'will', 'with', 'or', 'but', 'if', 'this', 'they'
        }
        
        return korean_stop_words | english_stop_words
    
    def _tokenize(self, text: str) -> List[str]:
        """
        텍스트를 토큰으로 분할합니다.
        
        Args:
            text (str): 분할할 텍스트
            
        Returns:
            List[str]: 토큰 목록
        """
        # 한글, 영문, 숫자만 남기고 소문자 변환
        text = re.sub(r'[^가-힣a-zA-Z0-9\s]', ' ', text.lower())
        
        # 공백으로 분할
        tokens = text.split()
        
        # 불용어 제거 및 길이 필터링 (2글자 이상)
        filtered_tokens = [
            token for token in tokens 
            if len(token) >= 2 and token not in self.stop_words
        ]
        
        return filtered_tokens
    
    def add_file(self, file_path: str, content: str, file_info: Dict[str, Any], pages_data: List[Dict[str, Any]] = None):
        """
        파일을 인덱스에 추가합니다.
        
        Args:
            file_path (str): 파일 경로
            content (str): 파일 내용 (전체)
            file_info (Dict[str, Any]): 파일 정보
            pages_data (List[Dict[str, Any]], optional): 페이지별 데이터 [{"page_num": 1, "content": "..."}, ...]
        """
        with self.lock:
            # 기존 인덱스에서 해당 파일 제거
            self.remove_file(file_path)
            
            # 파일 정보 저장
            self.file_info[file_path] = {
                **file_info,
                'indexed_time': datetime.now(),
                'content_preview': content[:200] if content else '',
                'full_content': content,  # 전체 내용도 저장 (검색용)
                'pages': pages_data if pages_data else None,  # 페이지별 데이터
            }
            
            # 파일명도 인덱싱에 포함
            filename = os.path.basename(file_path)
            all_content = f"{filename} {content}"
            
            # 텍스트 토큰화
            tokens = self._tokenize(all_content)
            
            # 역 인덱스 구축
            for token in set(tokens):  # 중복 제거
                self.index[token].add(file_path)
    
    def remove_file(self, file_path: str):
        """
        파일을 인덱스에서 제거합니다.
        
        Args:
            file_path (str): 제거할 파일 경로
        """
        with self.lock:
            # 파일 정보 제거
            if file_path in self.file_info:
                del self.file_info[file_path]
            
            # 인덱스에서 해당 파일 제거
            to_remove = []
            for token, file_paths in self.index.items():
                if file_path in file_paths:
                    file_paths.discard(file_path)
                    if not file_paths:  # 빈 집합이면 토큰 제거
                        to_remove.append(token)
            
            for token in to_remove:
                del self.index[token]
    
    def _search_keyword(self, query_tokens: List[str], enable_whitespace_agnostic: bool = True) -> set:
        """
        주어진 토큰들로 검색을 수행하고 매칭되는 파일들의 집합을 반환합니다.
        
        Args:
            query_tokens (List[str]): 검색할 토큰들
            enable_whitespace_agnostic (bool): 공백 무시 검색 활성화 여부
            
        Returns:
            set: 매칭되는 파일 경로들의 집합
        """
        # 각 토큰별로 매칭되는 파일 찾기
        token_results = []
        for token in query_tokens:
            matching_files = set()
            
            # 정확히 일치하는 토큰
            if token in self.index:
                matching_files.update(self.index[token])
            
            # 부분 일치하는 토큰 (접두사 매칭)
            for indexed_token in self.index:
                if indexed_token.startswith(token) or token in indexed_token:
                    matching_files.update(self.index[indexed_token])
            
            # 🆕 공백 무시 검색: full_content에서 직접 검색
            if enable_whitespace_agnostic:
                token_no_space = token.replace(' ', '').replace('\n', '').replace('\t', '')
                for file_path, file_info in self.file_info.items():
                    content = file_info.get('full_content', '').lower()
                    content_no_space = content.replace(' ', '').replace('\n', '').replace('\t', '')
                    
                    # 일반 검색 또는 공백 무시 검색
                    if token in content or token_no_space in content_no_space:
                        matching_files.add(file_path)
            
            token_results.append(matching_files)
        
        if not token_results:
            return set()
        
        # AND 연산 (모든 토큰이 포함된 파일)
        result_files = token_results[0]
        for token_result in token_results[1:]:
            result_files &= token_result
        
        return result_files
    
    def search(self, query: str, max_results: int = 50) -> List[Dict[str, Any]]:
        """
        검색 쿼리를 실행합니다.
        
        Args:
            query (str): 검색 쿼리 (쉼표로 분리된 다중 키워드 지원)
            max_results (int): 최대 결과 수
            
        Returns:
            List[Dict[str, Any]]: 검색 결과
        """
        with self.lock:
            if not query.strip():
                return []
            
            # 🆕 쉼표 기준 다중 키워드 분리 지원
            if ',' in query:
                # 쉼표로 분리된 키워드들 처리
                keywords = [kw.strip() for kw in query.split(',') if kw.strip()]
                if not keywords:
                    return []
                
                # 각 키워드별로 토큰화하고 모든 키워드가 포함된 파일만 찾기
                all_keyword_results = []
                for keyword in keywords:
                    keyword_tokens = self._tokenize(keyword)
                    keyword_files = self._search_keyword(keyword_tokens)
                    
                    if keyword_tokens:
                        all_keyword_results.append(keyword_files)
                
                if not all_keyword_results:
                    return []
                
                # 모든 키워드가 포함된 파일들만 교집합으로 찾기
                result_files = all_keyword_results[0]
                for keyword_files in all_keyword_results[1:]:
                    result_files &= keyword_files
                
                result_files = list(result_files)[:max_results]
                
                # 전체 쿼리 토큰화 (하이라이팅용)
                all_tokens = []
                for keyword in keywords:
                    all_tokens.extend(self._tokenize(keyword))
                query_tokens = all_tokens
            else:
                # 기존 단일 키워드 검색
                query_tokens = self._tokenize(query)
                if not query_tokens:
                    return []
                
                result_files = list(self._search_keyword(query_tokens))[:max_results]
            
            # 검색 결과 구성
            search_results = []
            for file_path in result_files:
                if file_path in self.file_info:
                    file_info = self.file_info[file_path]
                    
                    # 매칭된 컨텍스트 추출
                    content_preview = file_info.get('content_preview', '')
                    highlighted_preview = self._highlight_matches(content_preview, query_tokens)
                    
                    # 페이지 번호 찾기
                    matching_pages = self._find_matching_pages(file_info, query_tokens)
                    
                    result = {
                        'file_path': file_path,
                        'filename': os.path.basename(file_path),
                        'file_type': file_info.get('file_type', 'unknown'),
                        'file_size_mb': file_info.get('file_size_mb', 0),
                        'indexed_time': file_info.get('indexed_time'),
                        'preview': highlighted_preview,
                        'relevance_score': self._calculate_relevance(file_path, query_tokens),
                        'matching_pages': matching_pages  # 검색어가 있는 페이지 번호들
                    }
                    search_results.append(result)
            
            # 관련성 점수로 정렬
            search_results.sort(key=lambda x: x['relevance_score'], reverse=True)
            
            return search_results
    
    def _find_matching_pages(self, file_info: Dict[str, Any], query_tokens: List[str]) -> List[int]:
        """
        검색어가 포함된 페이지 번호들을 찾습니다.
        
        Args:
            file_info (Dict[str, Any]): 파일 정보
            query_tokens (List[str]): 검색 토큰들
            
        Returns:
            List[int]: 매칭된 페이지 번호 목록
        """
        pages_data = file_info.get('pages')
        if not pages_data:
            return []
        
        matching_pages = []
        for page in pages_data:
            page_num = page.get('page_num')
            content = page.get('content', '').lower()
            
            # 모든 검색 토큰이 이 페이지에 있는지 확인
            all_tokens_found = True
            for token in query_tokens:
                # 일반 검색
                if token.lower() in content:
                    continue
                # 공백 무시 검색
                token_no_space = token.replace(' ', '').replace('\n', '').replace('\t', '')
                content_no_space = content.replace(' ', '').replace('\n', '').replace('\t', '')
                if token_no_space in content_no_space:
                    continue
                # 토큰이 없으면 이 페이지는 매칭 안 됨
                all_tokens_found = False
                break
            
            if all_tokens_found and page_num != 0:  # Word 문서(page_num=0)는 제외
                matching_pages.append(page_num)
        
        return sorted(matching_pages)
    
    def _highlight_matches(self, text: str, query_tokens: List[str]) -> str:
        """
        텍스트에서 매칭된 부분을 하이라이트합니다.
        
        Args:
            text (str): 원본 텍스트
            query_tokens (List[str]): 검색 토큰들
            
        Returns:
            str: 하이라이트된 텍스트
        """
        highlighted = text
        
        for token in query_tokens:
            # 대소문자 구분 없이 매칭
            pattern = re.compile(re.escape(token), re.IGNORECASE)
            highlighted = pattern.sub(f"**{token}**", highlighted)
        
        return highlighted
    
    def _calculate_relevance(self, file_path: str, query_tokens: List[str]) -> float:
        """
        파일의 관련성 점수를 계산합니다.
        
        Args:
            file_path (str): 파일 경로
            query_tokens (List[str]): 검색 토큰들
            
        Returns:
            float: 관련성 점수
        """
        if file_path not in self.file_info:
            return 0.0
        
        score = 0.0
        filename = os.path.basename(file_path).lower()
        
        for token in query_tokens:
            # 파일명 매칭 보너스
            if token in filename:
                score += 2.0
            
            # 토큰 빈도 점수
            if token in self.index:
                # 드물게 나타나는 토큰일수록 높은 점수
                frequency = len(self.index[token])
                if frequency > 0:
                    score += 1.0 / frequency
        
        return score
    
    def get_statistics(self) -> Dict[str, Any]:
        """
        인덱스 통계 정보를 반환합니다.
        
        Returns:
            Dict[str, Any]: 통계 정보
        """
        with self.lock:
            return {
                'total_files': len(self.file_info),
                'total_tokens': len(self.index),
                'average_tokens_per_file': len(self.index) / max(len(self.file_info), 1),
                'file_types': self._get_file_type_distribution(),
            }
    
    def _get_file_type_distribution(self) -> Dict[str, int]:
        """파일 타입별 분포를 반환합니다."""
        distribution = defaultdict(int)
        for file_info in self.file_info.values():
            file_type = file_info.get('file_type', 'unknown')
            distribution[file_type] += 1
        return dict(distribution)


class SearchIndexer:
    """
    검색 인덱서 메인 클래스입니다.
    
    파일 시스템을 모니터링하고 자동으로 인덱싱을 수행합니다.
    JSON 캐싱 시스템으로 빠른 검색을 지원합니다.
    """
    
    def __init__(self):
        """SearchIndexer 인스턴스를 초기화합니다."""
        self.file_manager = FileManager()
        self.index = SearchIndex()
        self.indexing_thread = None
        self.stop_indexing = False
        self.indexed_paths = set()
        
        # [시작] JSON 캐싱 시스템 (사용자 요청)
        self.cache_directory = None
        self.cache_file_path = None
        self.metadata_file_path = None
    
    def index_directory(self, directory_path: str, recursive: bool = True, 
                       progress_callback=None):
        """
        디렉토리를 인덱싱합니다. (JSON 캐싱 시스템 포함)
        
        Args:
            directory_path (str): 인덱싱할 디렉토리 경로
            recursive (bool): 하위 디렉토리 포함 여부
            progress_callback: 진행 상태 콜백 함수
        """
        if not os.path.exists(directory_path):
            return
        
        # [시작] 캐시 디렉토리 설정 (사용자 요청: 동일 경로에 JSON 파일)
        self.set_cache_directory(directory_path)
        
        # [폴더] 캐시에서 기존 인덱스 로드 시도
        cache_loaded, files_to_reindex, new_files = self.load_index_from_cache(directory_path, recursive)
        
        print(f"[폴더] 디렉토리 인덱싱 시작: {directory_path}")
        if cache_loaded:
            print("[변환기] 캐시에서 기존 인덱스 로드됨. 변경된 파일만 처리합니다.")
        
        start_time = time.time()
        indexed_count = 0
        
        try:
            # 파일 목록 수집
            files_to_index = []
            
            if cache_loaded:
                # [시작] 캐시가 있을 때: 변경된 파일 + 새로운 파일만 처리
                files_to_index = files_to_reindex + new_files
                print(f"🎨 스마트 인덱싱: 변경된 파일 {len(files_to_reindex)}개 + 새로운 파일 {len(new_files)}개")
            else:
                # 💻 첫 인덱싱: 전체 디렉토리 스캔
                if recursive:
                    for root, dirs, files in os.walk(directory_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            if self.file_manager.is_supported_file(file_path):
                                # 엑셀 파일은 인덱싱에서 제외 (성능상 이유)
                                file_type = self.file_manager.get_file_type(file_path)
                                if file_type != 'excel':
                                    files_to_index.append(file_path)
                else:
                    for item in os.listdir(directory_path):
                        file_path = os.path.join(directory_path, item)
                        if os.path.isfile(file_path) and self.file_manager.is_supported_file(file_path):
                            # 엑셀 파일은 인덱싱에서 제외 (성능상 이유)
                            file_type = self.file_manager.get_file_type(file_path)
                            if file_type != 'excel':
                                files_to_index.append(file_path)
            
            total_files = len(files_to_index)
            if cache_loaded:
                print(f"[파일] 인덱싱 대상 파일: {total_files}개 (변경/신규 파일만)")
            else:
                print(f"[파일] 인덱싱 대상 파일: {total_files}개 (전체 파일)")
            
            # [변환기] 빠른 바이패스: 인덱싱할 파일이 없으면 스킵 (단, 캐시 업데이트는 필요)
            if total_files == 0:
                print("[완료] 변경된 파일이 없습니다. 인덱싱 완료!")
                # 삭제된 파일이 있다면 캐시 업데이트
                if cache_loaded:
                    self.save_index_to_cache()
                return
            
            # [시작] 멀티스레드 인덱싱 (사용자 요청: 3-4개 스레드로 속도 최대화)
            max_workers = min(4, max(1, len(files_to_index) // 10))  # 최적 스레드 수
            print(f"[변환기] {max_workers}개 스레드로 병렬 인덱싱 시작...")
            
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                # 파일들을 스레드에 분배
                futures = {
                    executor.submit(self._index_single_file, file_path): file_path 
                    for file_path in files_to_index
                }
                
                # 완료된 작업들 처리
                for i, future in enumerate(as_completed(futures)):
                    if self.stop_indexing:
                        break
                    
                    try:
                        result = future.result()
                        if result:  # 성공적으로 인덱싱됨
                            indexed_count += 1
                        
                        # 진행 상태 콜백
                        if progress_callback:
                            progress = (i + 1) / total_files * 100
                            file_path = futures[future]
                            progress_callback(file_path, progress)
                    
                    except Exception as e:
                        file_path = futures[future]
                        print(f"[오류] 파일 인덱싱 오류 ({file_path}): {e}")
            
            elapsed_time = time.time() - start_time
            if cache_loaded:
                print(f"[성공] 스마트 인덱싱 완료: {indexed_count}개 파일 처리, {elapsed_time:.2f}초 소요 (캐시 사용)")
            else:
                print(f"[성공] 전체 인덱싱 완료: {indexed_count}개 파일, {elapsed_time:.2f}초 소요")
            
            # [시작] 인덱싱 완료 후 JSON 캐시 저장 (사용자 요청)
            if indexed_count > 0 or cache_loaded:
                self.save_index_to_cache()  # 캐시가 있어도 업데이트
            
        except Exception as e:
            print(f"[오류] 디렉토리 인덱싱 오류: {e}")
    
    def search_files(self, query: str, max_results: int = 50) -> List[Dict[str, Any]]:
        """
        파일을 검색합니다. (JSON 캐시 우선, 폴백으로 메모리 인덱스)
        
        Args:
            query (str): 검색 쿼리
            max_results (int): 최대 결과 수
            
        Returns:
            List[Dict[str, Any]]: 검색 결과
        """
        # [시작] JSON 캐시에서 우선 검색 (사용자 요청: JSON에서 바로 검색)
        if self.cache_file_path and os.path.exists(self.cache_file_path):
            return self.search_files_from_json(query, max_results)
        
        # 폴백: 메모리 인덱스에서 검색
        print("[경고] JSON 캐시 없음. 메모리 인덱스에서 검색...")
        return self.index.search(query, max_results)
    
    def add_file_to_index(self, file_path: str):
        """
        개별 파일을 인덱스에 추가합니다.
        
        Args:
            file_path (str): 추가할 파일 경로
        """
        try:
            if self.file_manager.is_supported_file(file_path):
                # 엑셀 파일은 인덱싱에서 제외 (성능상 이유)
                file_type = self.file_manager.get_file_type(file_path)
                if file_type == 'excel':
                    print(f"[INFO] 엑셀 파일은 인덱싱에서 제외됨: {file_path}")
                    return
                
                file_info = self.file_manager.get_file_info(file_path)
                
                if file_info.get('supported', False):
                    content = self.file_manager.extract_text(file_path)
                    
                    # 페이지별 데이터 추출 (PDF/PowerPoint만)
                    pages_data = None
                    if file_type == 'pdf':
                        pdf_handler = self.file_manager.handlers.get('pdf')
                        if pdf_handler:
                            pages_data = pdf_handler.extract_text_by_pages(file_path)
                    elif file_type == 'powerpoint':
                        ppt_handler = self.file_manager.handlers.get('powerpoint')
                        if ppt_handler:
                            pages_data = ppt_handler.extract_text_by_slides(file_path)
                    elif file_type == 'word':
                        # Word는 페이지 구분 어려움 - 메타데이터로 표시
                        pages_data = [{"page_num": 0, "content": "(Word 문서는 페이지 구분이 어렵습니다)"}]
                    
                    self.index.add_file(file_path, content, file_info, pages_data)
                    self.indexed_paths.add(file_path)
                    print(f"[OK] 파일 인덱싱 완료: {file_path}")
        
        except Exception as e:
            print(f"[ERROR] 파일 인덱싱 오류 ({file_path}): {e}")
    
    def remove_file_from_index(self, file_path: str):
        """
        파일을 인덱스에서 제거합니다.
        
        Args:
            file_path (str): 제거할 파일 경로
        """
        self.index.remove_file(file_path)
        self.indexed_paths.discard(file_path)
        print(f"[삭제] 파일 인덱스 제거: {file_path}")
    
    def update_file_in_index(self, file_path: str):
        """
        파일 인덱스를 업데이트합니다.
        
        Args:
            file_path (str): 업데이트할 파일 경로
        """
        if file_path in self.indexed_paths:
            self.remove_file_from_index(file_path)
        
        self.add_file_to_index(file_path)
    
    def get_index_statistics(self) -> Dict[str, Any]:
        """
        인덱스 통계를 반환합니다.
        
        Returns:
            Dict[str, Any]: 통계 정보
        """
        stats = self.index.get_statistics()
        stats['indexed_paths_count'] = len(self.indexed_paths)
        return stats
    
    def clear_index(self):
        """인덱스를 초기화합니다."""
        self.index = SearchIndex()
        self.indexed_paths.clear()
        print("🧹 검색 인덱스가 초기화되었습니다.")
    
    def stop_indexing_process(self):
        """인덱싱 프로세스를 중단합니다."""
        self.stop_indexing = True
    
    def set_cache_directory(self, directory_path: str):
        """
        캐시 디렉토리를 설정합니다. (사용자 제안: 동일 경로에 JSON 파일 저장)
        
        Args:
            directory_path (str): 검색 대상 디렉토리 경로
        """
        self.cache_directory = directory_path
        self.cache_file_path = os.path.join(directory_path, ".file_index.json")
        self.metadata_file_path = os.path.join(directory_path, ".index_metadata.json")
        print(f"[폴더] 캐시 설정: {self.cache_file_path}")
    
    def _get_file_hash(self, file_path: str) -> str:
        """
        파일의 해시값을 계산합니다. (수정 시간 + 크기 기반)
        
        Args:
            file_path (str): 파일 경로
            
        Returns:
            str: 파일 해시값
        """
        try:
            stat = os.stat(file_path)
            # 수정 시간 + 크기로 해시 생성
            hash_input = f"{stat.st_mtime}_{stat.st_size}_{file_path}"
            return hashlib.md5(hash_input.encode('utf-8')).hexdigest()
        except:
            return ""
    
    def save_index_to_cache(self):
        """
        인덱스를 JSON 파일에 저장합니다. (사용자 요청: JSON 파일로 캐싱)
        """
        if not self.cache_file_path or not self.cache_directory:
            print("[경고] 캐시 경로가 설정되지 않음")
            return
        
        try:
            print("[저장] 인덱스를 JSON 파일에 저장 중...")
            
            # 인덱스 데이터 구성 (사용자 제안 방식)
            cache_data = {
                "files": {},
                "last_indexed": datetime.now().isoformat(),
                "total_files": len(self.indexed_paths),
                "index_version": "1.0"
            }
            
            # 파일별 정보 저장 (파일명 + 내용)
            for file_path in self.indexed_paths:
                if file_path in self.index.file_info:
                    file_info = self.index.file_info[file_path]
                    relative_path = os.path.relpath(file_path, str(self.cache_directory))
                    
                    cache_data["files"][relative_path] = {
                        "content": file_info.get('full_content', ''),
                        "title": os.path.basename(file_path),
                        "size": file_info.get('file_size_mb', 0),
                        "modified": file_info.get('indexed_time', datetime.now()).isoformat(),
                        "type": file_info.get('file_type', 'unknown'),
                        "file_hash": self._get_file_hash(file_path),
                        "full_path": file_path
                    }
            
            # JSON 파일로 저장
            with open(str(self.cache_file_path), 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
            
            # 메타데이터 저장
            metadata = {
                "cache_created": datetime.now().isoformat(),
                "indexed_directory": str(self.cache_directory),
                "total_files": len(self.indexed_paths),
                "cache_file_size": os.path.getsize(str(self.cache_file_path))
            }
            
            with open(str(self.metadata_file_path), 'w', encoding='utf-8') as f:
                json.dump(metadata, f, ensure_ascii=False, indent=2)
            
            print(f"[성공] 인덱스 캐시 저장 완료: {len(self.indexed_paths)}개 파일")
            
        except Exception as e:
            print(f"[오류] 인덱스 캐시 저장 실패: {e}")
    
    def load_index_from_cache(self, directory_path: str = None, recursive: bool = True) -> Tuple[bool, List[str], List[str]]:
        """
        JSON 파일에서 인덱스를 로드합니다. (사용자 요청: JSON에서 빠른 검색)
        
        Args:
            directory_path (str): 비교할 디렉토리 경로 (새 파일 감지용)
            recursive (bool): 하위 디렉토리 포함 여부
        
        Returns:
            tuple: (로드 성공 여부, 변경된 파일 리스트, 새로운 파일 리스트)
        """
        if not self.cache_file_path or not os.path.exists(self.cache_file_path):
            print("[파일] 캐시 파일이 없습니다. 새로 인덱싱이 필요합니다.")
            return False, [], []
        
        try:
            print("[폴더] JSON 캐시에서 인덱스 로드 중...")
            
            with open(str(self.cache_file_path), 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
            
            # 캐시 버전 체크
            if cache_data.get("index_version") != "1.0":
                print("[경고] 캐시 버전 불일치. 새로 인덱싱이 필요합니다.")
                return False, [], []
            
            # 파일 변경 사항 체크 (스마트 재인덱싱)
            files_to_reindex = []
            valid_files = 0
            
            for relative_path, file_data in cache_data["files"].items():
                full_path = file_data.get("full_path")
                if not full_path or not os.path.exists(full_path):
                    continue
                
                # 파일 해시 체크로 변경 감지
                current_hash = self._get_file_hash(full_path)
                cached_hash = file_data.get("file_hash", "")
                
                if current_hash != cached_hash:
                    files_to_reindex.append(full_path)
                else:
                    # 변경되지 않은 파일은 캐시에서 복원
                    file_info = {
                        'file_type': file_data.get('type', 'unknown'),
                        'file_size_mb': file_data.get('size', 0),
                        'indexed_time': datetime.fromisoformat(file_data.get('modified', datetime.now().isoformat())),
                        'content_preview': file_data.get('content', ''),
                        'supported': True
                    }
                    
                    # 인덱스에 추가
                    content = file_data.get('content', '')
                    self.index.add_file(full_path, content, file_info)
                    self.indexed_paths.add(full_path)
                    valid_files += 1
            
            print(f"[성공] 캐시에서 {valid_files}개 파일 로드 완료")
            
            # 새로운 파일 및 삭제된 파일 감지 (현재 디렉토리와 캐시 비교)
            new_files = []
            deleted_files = []
            
            if directory_path:
                # 현재 디렉토리의 지원 파일들 수집 (recursive 플래그 준수)
                current_files = set()
                
                if recursive:
                    for root, dirs, files in os.walk(directory_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            if self.file_manager.is_supported_file(file_path):
                                file_type = self.file_manager.get_file_type(file_path)
                                if file_type != 'excel':  # 엑셀 파일 제외
                                    # 🔧 경로 정규화: 일관성 있는 비교를 위해
                                    normalized_path = os.path.normcase(os.path.normpath(os.path.realpath(file_path)))
                                    current_files.add(normalized_path)
                else:
                    for item in os.listdir(directory_path):
                        file_path = os.path.join(directory_path, item)
                        if os.path.isfile(file_path) and self.file_manager.is_supported_file(file_path):
                            file_type = self.file_manager.get_file_type(file_path)
                            if file_type != 'excel':  # 엑셀 파일 제외
                                # 🔧 경로 정규화: 일관성 있는 비교를 위해
                                normalized_path = os.path.normcase(os.path.normpath(os.path.realpath(file_path)))
                                current_files.add(normalized_path)
                
                # 🚨 중요: 삭제 감지 범위를 current_files와 동일하게 제한
                # recursive=False일 때 하위폴더 파일을 "삭제됨"으로 잘못 판단하는 버그 방지
                
                # 경로 정규화 (대소문자, 구분자, 심볼릭링크 처리)
                normalized_directory = os.path.normcase(os.path.normpath(os.path.realpath(directory_path)))
                
                if recursive:
                    # recursive=True: 모든 캐시된 파일 고려 (정규화)
                    cached_files = set()
                    for file_data in cache_data["files"].values():
                        cached_path = file_data.get("full_path")
                        if cached_path:
                            # 🔧 캐시된 경로도 정규화: current_files와 일관성 유지
                            normalized_cached = os.path.normcase(os.path.normpath(os.path.realpath(cached_path)))
                            cached_files.add(normalized_cached)
                else:
                    # recursive=False: 현재 폴더의 캐시된 파일만 고려 (정규화된 경로 비교)
                    cached_files = set()
                    for file_data in cache_data["files"].values():
                        cached_path = file_data.get("full_path")
                        if cached_path:
                            # 🔧 캐시된 경로 정규화
                            normalized_cached = os.path.normcase(os.path.normpath(os.path.realpath(cached_path)))
                            cached_parent = os.path.normcase(os.path.normpath(os.path.realpath(os.path.dirname(cached_path))))
                            if cached_parent == normalized_directory:
                                cached_files.add(normalized_cached)
                
                # 새로운 파일 = 현재 파일 - 캐시된 파일 (정규화된 경로로 정확한 비교)
                new_files_normalized = list(current_files - cached_files)
                
                # 삭제된 파일 = 캐시된 파일 - 현재 파일 (정규화된 경로로 정확한 비교)
                deleted_files_normalized = list(cached_files - current_files)
                
                # [변경] 인덱싱을 위해 원본 절대 경로로 복원 (정규화되지 않은 원본 경로 사용)
                # new_files: 정규화된 경로에서 원본 절대 경로 매핑
                normalized_to_original = {}
                if recursive:
                    for root, dirs, files in os.walk(directory_path):
                        for file in files:
                            original_path = os.path.join(root, file)
                            if self.file_manager.is_supported_file(original_path):
                                file_type = self.file_manager.get_file_type(original_path)
                                if file_type != 'excel':
                                    normalized = os.path.normcase(os.path.normpath(os.path.realpath(original_path)))
                                    normalized_to_original[normalized] = original_path
                else:
                    for item in os.listdir(directory_path):
                        original_path = os.path.join(directory_path, item)
                        if os.path.isfile(original_path) and self.file_manager.is_supported_file(original_path):
                            file_type = self.file_manager.get_file_type(original_path)
                            if file_type != 'excel':
                                normalized = os.path.normcase(os.path.normpath(os.path.realpath(original_path)))
                                normalized_to_original[normalized] = original_path
                
                # 원본 경로로 복원
                new_files = [normalized_to_original.get(norm_path, norm_path) for norm_path in new_files_normalized]
                deleted_files = deleted_files_normalized  # 삭제된 파일은 정규화된 경로 사용
                
                # 삭제된 파일을 인덱스에서 제거
                for deleted_file in deleted_files:
                    if deleted_file in self.indexed_paths:
                        self.remove_file_from_index(deleted_file)
            
            if files_to_reindex:
                print(f"[변경] 변경된 파일 {len(files_to_reindex)}개 재인덱싱 필요")
            
            if new_files:
                print(f"[파일] 새로운 파일 {len(new_files)}개 발견")
            
            if deleted_files:
                print(f"[삭제] 삭제된 파일 {len(deleted_files)}개 제거")
            
            return True, files_to_reindex, new_files
            
        except Exception as e:
            print(f"[오류] 캐시 로드 실패: {e}")
            return False, [], []
    
    def get_cache_statistics(self) -> Dict[str, Any]:
        """
        캐시 통계 정보를 반환합니다.
        
        Returns:
            Dict[str, Any]: 캐시 통계
        """
        from typing import Dict, Any, Union
        stats: Dict[str, Union[bool, float, str]] = {"cache_available": False}
        
        if self.cache_file_path and os.path.exists(str(self.cache_file_path)):
            try:
                cache_size = os.path.getsize(str(self.cache_file_path))
                cache_modified = datetime.fromtimestamp(os.path.getmtime(str(self.cache_file_path)))
                
                stats["cache_available"] = True
                stats["cache_size_mb"] = float(cache_size) / (1024.0 * 1024.0)  # 명시적 float 변환
                stats["cache_modified"] = cache_modified.isoformat()
                stats["cache_file"] = str(self.cache_file_path)
                
            except:
                pass
        
        return stats
    
    def _index_single_file(self, file_path: str) -> bool:
        """
        단일 파일을 인덱싱합니다. (멀티스레드에서 사용)
        
        Args:
            file_path (str): 인덱싱할 파일 경로
            
        Returns:
            bool: 인덱싱 성공 여부
        """
        try:
            # 파일 정보 조회
            file_info = self.file_manager.get_file_info(file_path)
            
            if file_info.get('supported', False):
                # 텍스트 추출
                content = self.file_manager.extract_text(file_path)
                
                # 스레드 안전하게 인덱스에 추가
                self.index.add_file(file_path, content, file_info)
                self.indexed_paths.add(file_path)
                
                return True
            
            return False
            
        except Exception as e:
            print(f"[오류] 파일 인덱싱 오류 ({file_path}): {e}")
            return False
    
    def search_files_from_json(self, query: str, max_results: int = 50) -> List[Dict[str, Any]]:
        """
        JSON 캐시에서 직접 검색합니다. (사용자 요청: JSON에서 바로 빠른 검색)
        
        Args:
            query (str): 검색 쿼리
            max_results (int): 최대 결과 수
            
        Returns:
            List[Dict[str, Any]]: 검색 결과
        """
        if not self.cache_file_path or not os.path.exists(self.cache_file_path):
            print("[파일] JSON 캐시 파일이 없습니다. 인덱싱을 먼저 실행하세요.")
            return []
        
        try:
            print(f"🔍 JSON에서 '{query}' 검색 중...")
            
            with open(str(self.cache_file_path), 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
            
            results = []
            
            # 🆕 다중 키워드 지원 추가
            if ',' in query:
                keywords = [kw.strip().lower() for kw in query.split(',') if kw.strip()]
            else:
                keywords = [query.lower()]
            
            # 🆕 공백 제거 버전 키워드 생성 (띄어쓰기 무시 검색용)
            keywords_no_space = [kw.replace(' ', '').replace('\n', '').replace('\t', '') for kw in keywords]
            
            # 파일별로 검색 수행
            for relative_path, file_data in cache_data.get("files", {}).items():
                full_path = file_data.get("full_path", "")
                if not os.path.exists(full_path):
                    continue
                
                # 파일명 + 내용에서 검색
                title = file_data.get("title", "").lower()
                content = file_data.get("content", "").lower()
                search_text = f"{title} {content}"
                
                # 🆕 공백 제거 버전 텍스트 (띄어쓰기 무시 검색용)
                title_no_space = title.replace(' ', '').replace('\n', '').replace('\t', '')
                content_no_space = content.replace(' ', '').replace('\n', '').replace('\t', '')
                search_text_no_space = f"{title_no_space} {content_no_space}"
                
                # 🆕 모든 키워드가 포함되어야 함 (AND 검색)
                all_keywords_found = True
                filename_matches = 0
                content_matches = 0
                
                for i, keyword in enumerate(keywords):
                    keyword_no_space = keywords_no_space[i]
                    
                    # 일반 검색 + 공백 무시 검색
                    found_in_text = keyword in search_text or keyword_no_space in search_text_no_space
                    
                    if not found_in_text:
                        all_keywords_found = False
                        break
                    
                    # 개별 키워드별 매칭 체크 (일반 + 공백 무시)
                    if keyword in title or keyword_no_space in title_no_space:
                        filename_matches += 1
                    if keyword in content or keyword_no_space in content_no_space:
                        content_matches += 1
                
                # 매칭 체크
                filename_match = filename_matches > 0
                content_match = content_matches > 0
                
                if all_keywords_found and (filename_match or content_match):
                    # 관련성 점수 계산
                    relevance_score = 0.0
                    relevance_score += filename_matches * 2.0  # 파일명 매칭 키워드별 점수
                    relevance_score += content_matches * 1.0   # 내용 매칭 키워드별 점수
                    
                    # 매칭된 컨텍스트 추출
                    preview = self._extract_context_from_content(
                        file_data.get("content", ""), query
                    )
                    
                    result = {
                        'file_path': full_path,
                        'filename': file_data.get("title", ""),
                        'file_type': file_data.get("type", "unknown"),
                        'file_size_mb': file_data.get("size", 0),
                        'indexed_time': file_data.get("modified", ""),
                        'preview': preview,
                        'relevance_score': relevance_score
                    }
                    results.append(result)
            
            # 관련성 점수로 정렬
            results.sort(key=lambda x: x['relevance_score'], reverse=True)
            
            print(f"[성공] JSON 검색 완료: {len(results)}개 결과")
            return results[:max_results]
            
        except Exception as e:
            print(f"[오류] JSON 검색 실패: {e}")
            return []
    
    def search_files_by_filename_from_json(self, query: str, max_results: int = 50) -> List[Dict[str, Any]]:
        """
        JSON 캐시에서 파일명으로만 검색합니다. (초고속)
        
        Args:
            query (str): 검색 쿼리
            max_results (int): 최대 결과 수
            
        Returns:
            List[Dict[str, Any]]: 검색 결과
        """
        if not self.cache_file_path or not os.path.exists(self.cache_file_path):
            return []
        
        try:
            with open(str(self.cache_file_path), 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
            
            results = []
            query_lower = query.lower()
            
            # 파일명에서만 검색 (매우 빠름)
            for relative_path, file_data in cache_data.get("files", {}).items():
                full_path = file_data.get("full_path", "")
                if not os.path.exists(full_path):
                    continue
                
                title = file_data.get("title", "").lower()
                filename_without_ext = os.path.splitext(title)[0].lower()
                
                if query_lower in filename_without_ext:
                    relevance_score = 1.0
                    if filename_without_ext.startswith(query_lower):
                        relevance_score = 2.0  # 시작하는 경우 더 높은 점수
                    
                    result = {
                        'file_path': full_path,
                        'filename': file_data.get("title", ""),
                        'file_type': file_data.get("type", "unknown"),
                        'file_size_mb': file_data.get("size", 0),
                        'indexed_time': file_data.get("modified", ""),
                        'preview': f"파일명 매칭: {file_data.get('title', '')}",
                        'relevance_score': relevance_score
                    }
                    results.append(result)
            
            results.sort(key=lambda x: x['relevance_score'], reverse=True)
            return results[:max_results]
            
        except Exception as e:
            print(f"[오류] JSON 파일명 검색 실패: {e}")
            return []
    
    def _extract_context_from_content(self, content: str, query: str, context_length: int = 150) -> str:
        """
        검색어 주변의 컨텍스트를 추출합니다.
        
        Args:
            content (str): 원본 내용
            query (str): 검색어
            context_length (int): 컨텍스트 길이
            
        Returns:
            str: 하이라이트된 컨텍스트
        """
        if not content or not query:
            return content[:context_length] if content else ""
        
        content_lower = content.lower()
        query_lower = query.lower()
        
        # 첫 번째 매칭 위치 찾기
        match_pos = content_lower.find(query_lower)
        if match_pos == -1:
            return content[:context_length]
        
        # 컨텍스트 시작/끝 위치 계산
        start = max(0, match_pos - context_length // 2)
        end = min(len(content), match_pos + len(query) + context_length // 2)
        
        context = content[start:end]
        
        # 검색어 하이라이트 (간단한 방식)
        highlighted = context.replace(
            content[match_pos:match_pos + len(query)], 
            f"**{content[match_pos:match_pos + len(query)]}**"
        )
        
        # 앞뒤 생략 표시
        if start > 0:
            highlighted = "..." + highlighted
        if end < len(content):
            highlighted = highlighted + "..."
        
        return highlighted