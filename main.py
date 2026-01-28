import requests
import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET
from urllib.parse import quote


class G2BAPIClient:
    def __init__(self, service_key):
        """
        조달청 나라장터 OpenAPI 클라이언트 초기화

        Args:
            service_key: 공공데이터포털에서 발급받은 서비스키
        """
        self.base_url = "http://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoThngPPSSrch"
        self.service_key = service_key

    def fetch_bid_notices(self, search_params):
        """
        입찰공고 정보 조회

        Args:
            search_params: 검색 파라미터 딕셔너리
                - numOfRows: 한 페이지 결과 수
                - pageNo: 페이지 번호
                - inqryDiv: 조회구분 (1:공고게시일시, 2:개찰일시)
                - inqryBgnDt: 조회시작일시 (YYYYMMDDHHMM)
                - inqryEndDt: 조회종료일시 (YYYYMMDDHHMM)
                - bidNtceNm: 입찰공고명 (선택)

        Returns:
            파싱된 입찰공고 데이터 리스트
        """
        params = {
            'ServiceKey': self.service_key,
            'numOfRows': search_params.get('numOfRows', 100),
            'pageNo': search_params.get('pageNo', 10),
            'inqryDiv': search_params.get('inqryDiv', '1'),
            'inqryBgnDt': search_params['inqryBgnDt'],
            'inqryEndDt': search_params['inqryEndDt']
        }

        # 선택적 파라미터 추가
        if 'bidNtceNm' in search_params and search_params['bidNtceNm']:
            params['bidNtceNm'] = search_params['bidNtceNm']

        try:
            response = requests.get(self.base_url, params=params, timeout=30)
            response.raise_for_status()

            # XML 파싱
            root = ET.fromstring(response.content)

            # 결과 코드 확인
            result_code = root.find('.//resultCode').text if root.find('.//resultCode') is not None else None
            result_msg = root.find('.//resultMsg').text if root.find('.//resultMsg') is not None else None

            if result_code != '00':
                print(f"API 오류: {result_msg}")
                return []

            # 데이터 파싱
            items = root.findall('.//item')
            return self._parse_items(items)

        except requests.exceptions.RequestException as e:
            print(f"API 호출 오류: {e}")
            return []
        except ET.ParseError as e:
            print(f"XML 파싱 오류: {e}")
            return []

    def _parse_items(self, items):
        """XML item 요소를 파싱하여 딕셔너리 리스트로 변환"""
        result = []

        for item in items:
            data = {
                'bidNtceNo': self._get_text(item, 'bidNtceNo'),
                'rgstTyNm': self._get_text(item, 'rgstTyNm'),
                'ntceKindNm': self._get_text(item, 'ntceKindNm'),
                'bidNtceDt': self._get_text(item, 'bidNtceDt'),
                'bidNtceNm': self._get_text(item, 'bidNtceNm'),
                'ntceInsttCd': self._get_text(item, 'ntceInsttCd'),
                'ntceInsttNm': self._get_text(item, 'ntceInsttNm'),
                'dminsttCd': self._get_text(item, 'dminsttCd'),
                'dminsttNm': self._get_text(item, 'dminsttNm'),
                'ntceInsttOfclNm': self._get_text(item, 'ntceInsttOfclNm'),
                'ntceInsttOfclTelNo': self._get_text(item, 'ntceInsttOfclTelNo'),
                'ntceInsttOfclEmailAdrs': self._get_text(item, 'ntceInsttOfclEmailAdrs'),
                'exctvNm': self._get_text(item, 'exctvNm'),
                'bidQlfctRgstDt': self._get_text(item, 'bidQlfctRgstDt'),
                'bidBeginDt': self._get_text(item, 'bidBeginDt'),
                'bidClseDt': self._get_text(item, 'bidClseDt'),
                'opengDt': self._get_text(item, 'opengDt')
            }
            result.append(data)

        return result

    def _get_text(self, item, tag_name):
        """XML 요소에서 텍스트 추출 (없으면 빈 문자열 반환)"""
        element = item.find(tag_name)
        return element.text if element is not None and element.text else ''

    def fetch_all_pages(self, search_params):
        """
        모든 페이지의 데이터를 가져오기

        Args:
            search_params: 검색 파라미터

        Returns:
            전체 입찰공고 데이터 리스트
        """
        all_data = []
        page_no = 1
        num_of_rows = search_params.get('numOfRows', 100)

        while True:
            print(f"페이지 {page_no} 조회 중...")
            search_params['pageNo'] = page_no
            search_params['numOfRows'] = num_of_rows

            data = self.fetch_bid_notices(search_params)

            if not data:
                break

            all_data.extend(data)

            # 마지막 페이지인지 확인 (가져온 데이터가 요청한 개수보다 적으면 마지막 페이지)
            if len(data) < num_of_rows:
                break

            page_no += 1

        print(f"총 {len(all_data)}건의 데이터를 가져왔습니다.")
        return all_data


def save_to_excel(data, filename=None):
    """
    데이터를 엑셀 파일로 저장

    Args:
        data: 입찰공고 데이터 리스트
        filename: 저장할 파일명 (기본값: 현재시각_입찰공고.xlsx)
    """
    if not data:
        print("저장할 데이터가 없습니다.")
        return

    # DataFrame 생성
    df = pd.DataFrame(data)

    # 컬럼명 한글화
    df.columns = [
        '입찰공고번호',
        '등록유형명',
        '공고종류명',
        '입찰공고일시',
        '입찰공고명',
        '공고기관코드',
        '공고기관명',
        '수요기관코드',
        '수요기관명',
        '공고기관담당자명',
        '공고기관담당자전화번호',
        '공고기관담당자이메일주소',
        '집행관명',
        '입찰참가자격등록마감일시',
        '입찰개시일시',
        '입찰마감일시',
        '개찰일시'
    ]

    # 파일명 생성
    if filename is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'{timestamp}_입찰공고.xlsx'

    # 엑셀 저장
    df.to_excel(filename, index=False, engine='openpyxl')
    print(f"엑셀 파일이 저장되었습니다: {filename}")


def main():
    """메인 실행 함수"""

    # 서비스키 설정 (공공데이터포털에서 발급받은 키를 입력하세요)
    SERVICE_KEY = ""

    # API 클라이언트 생성
    client = G2BAPIClient(SERVICE_KEY)

    # 검색 조건 설정
    search_params = {
        'inqryDiv': '1',  # 1:공고게시일시, 2:개찰일시
        'inqryBgnDt': '202601010000',  # 조회시작일시 (YYYYMMDDHHMM)
        'inqryEndDt': '202601312359',  # 조회종료일시 (YYYYMMDDHHMM)
        'numOfRows': 200,  # 한 페이지당 결과 수
        'bidNtceNm': '그래픽카드'  # 입찰공고명 (선택사항)
    }

    # 데이터 조회
    print("입찰공고 데이터 조회 시작...")
    bid_data = client.fetch_all_pages(search_params)

    # 엑셀 저장
    if bid_data:
        save_to_excel(bid_data)

        # 담당자 이메일만 추출하여 별도 출력
        print("\n=== 담당자 이메일 목록 ===")
        emails = [item['ntceInsttOfclEmailAdrs'] for item in bid_data if item['ntceInsttOfclEmailAdrs']]
        unique_emails = list(set(emails))  # 중복 제거
        for email in unique_emails:
            print(email)
    else:
        print("조회된 데이터가 없습니다.")


if __name__ == "__main__":
    main()