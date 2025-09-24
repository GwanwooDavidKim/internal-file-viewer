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
from datetime import datetime
from typing import Dict, List, Any, Optional, Set, Tuple
from collections import defaultdict
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
    
    def add_file(self, file_path: str, content: str, file_info: Dict[str, Any]):
        """
        파일을 인덱스에 추가합니다.
        
        Args:
            file_path (str): 파일 경로
            content (str): 파일 내용
            file_info (Dict[str, Any]): 파일 정보
        """
        with self.lock:
            # 기존 인덱스에서 해당 파일 제거
            self.remove_file(file_path)
            
            # 파일 정보 저장
            self.file_info[file_path] = {
                **file_info,
                'indexed_time': datetime.now(),
                'content_preview': content[:200] if content else '',
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
    
    def search(self, query: str, max_results: int = 50) -> List[Dict[str, Any]]:
        """
        검색 쿼리를 실행합니다.
        
        Args:
            query (str): 검색 쿼리
            max_results (int): 최대 결과 수
            
        Returns:
            List[Dict[str, Any]]: 검색 결과
        """
        with self.lock:
            if not query.strip():
                return []
            
            # 쿼리 토큰화
            query_tokens = self._tokenize(query)
            if not query_tokens:
                return []
            
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
                
                token_results.append(matching_files)
            
            if not token_results:
                return []
            
            # AND 연산 (모든 토큰이 포함된 파일)
            result_files = token_results[0]
            for token_result in token_results[1:]:
                result_files &= token_result
            
            # 결과가 적으면 OR 연산도 포함
            if len(result_files) < max_results // 2:
                or_results = set()
                for token_result in token_results:
                    or_results |= token_result
                
                # AND 결과를 우선하고 OR 결과를 추가
                result_files = list(result_files) + list(or_results - result_files)
            else:
                result_files = list(result_files)
            
            # 결과 제한
            result_files = result_files[:max_results]
            
            # 검색 결과 구성
            search_results = []
            for file_path in result_files:
                if file_path in self.file_info:
                    file_info = self.file_info[file_path]
                    
                    # 매칭된 컨텍스트 추출
                    content_preview = file_info.get('content_preview', '')
                    highlighted_preview = self._highlight_matches(content_preview, query_tokens)
                    
                    result = {
                        'file_path': file_path,
                        'filename': os.path.basename(file_path),
                        'file_type': file_info.get('file_type', 'unknown'),
                        'file_size_mb': file_info.get('file_size_mb', 0),
                        'indexed_time': file_info.get('indexed_time'),
                        'preview': highlighted_preview,
                        'relevance_score': self._calculate_relevance(file_path, query_tokens)
                    }
                    search_results.append(result)
            
            # 관련성 점수로 정렬
            search_results.sort(key=lambda x: x['relevance_score'], reverse=True)
            
            return search_results
    
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
    """
    
    def __init__(self):
        """SearchIndexer 인스턴스를 초기화합니다."""
        self.file_manager = FileManager()
        self.index = SearchIndex()
        self.indexing_thread = None
        self.stop_indexing = False
        self.indexed_paths = set()
    
    def index_directory(self, directory_path: str, recursive: bool = True, 
                       progress_callback=None):
        """
        디렉토리를 인덱싱합니다.
        
        Args:
            directory_path (str): 인덱싱할 디렉토리 경로
            recursive (bool): 하위 디렉토리 포함 여부
            progress_callback: 진행 상태 콜백 함수
        """
        if not os.path.exists(directory_path):
            return
        
        print(f"📂 디렉토리 인덱싱 시작: {directory_path}")
        start_time = time.time()
        indexed_count = 0
        
        try:
            # 파일 목록 수집
            files_to_index = []
            
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
            print(f"📄 인덱싱 대상 파일: {total_files}개")
            
            # 파일별 인덱싱
            for i, file_path in enumerate(files_to_index):
                if self.stop_indexing:
                    break
                
                try:
                    # 파일 정보 조회
                    file_info = self.file_manager.get_file_info(file_path)
                    
                    if file_info.get('supported', False):
                        # 텍스트 추출
                        content = self.file_manager.extract_text(file_path)
                        
                        # 인덱스에 추가
                        self.index.add_file(file_path, content, file_info)
                        self.indexed_paths.add(file_path)
                        indexed_count += 1
                        
                        # 진행 상태 콜백
                        if progress_callback:
                            progress = (i + 1) / total_files * 100
                            progress_callback(file_path, progress)
                
                except Exception as e:
                    print(f"❌ 파일 인덱싱 오류 ({file_path}): {e}")
            
            elapsed_time = time.time() - start_time
            print(f"✅ 인덱싱 완료: {indexed_count}개 파일, {elapsed_time:.2f}초 소요")
            
        except Exception as e:
            print(f"❌ 디렉토리 인덱싱 오류: {e}")
    
    def search_files(self, query: str, max_results: int = 50) -> List[Dict[str, Any]]:
        """
        파일을 검색합니다.
        
        Args:
            query (str): 검색 쿼리
            max_results (int): 최대 결과 수
            
        Returns:
            List[Dict[str, Any]]: 검색 결과
        """
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
                    print(f"⚠️ 엑셀 파일은 인덱싱에서 제외됨: {file_path}")
                    return
                
                file_info = self.file_manager.get_file_info(file_path)
                
                if file_info.get('supported', False):
                    content = self.file_manager.extract_text(file_path)
                    self.index.add_file(file_path, content, file_info)
                    self.indexed_paths.add(file_path)
                    print(f"✅ 파일 인덱싱 완료: {file_path}")
        
        except Exception as e:
            print(f"❌ 파일 인덱싱 오류 ({file_path}): {e}")
    
    def remove_file_from_index(self, file_path: str):
        """
        파일을 인덱스에서 제거합니다.
        
        Args:
            file_path (str): 제거할 파일 경로
        """
        self.index.remove_file(file_path)
        self.indexed_paths.discard(file_path)
        print(f"🗑️ 파일 인덱스 제거: {file_path}")
    
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