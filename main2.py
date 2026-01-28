import pandas as pd
from datetime import datetime


def create_unique_contact_list(input_file, output_file=None):
    """
    입찰공고 엑셀 파일에서 수요기관 담당자 정보를 추출하여 중복 제거 후 새 엑셀 생성

    Args:
        input_file: 입력 엑셀 파일 경로
        output_file: 출력 엑셀 파일 경로 (기본값: 자동생성)
    """
    try:
        # 엑셀 파일 읽기
        print(f"'{input_file}' 파일을 읽는 중...")
        df = pd.read_excel(input_file)

        # 필요한 컬럼만 선택
        required_columns = [
            '수요기관코드',
            '수요기관명',
            '공고기관담당자명',
            '공고기관담당자전화번호',
            '공고기관담당자이메일주소'
        ]

        # 컬럼 존재 여부 확인
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"오류: 다음 컬럼이 파일에 없습니다: {missing_columns}")
            return

        # 필요한 컬럼만 추출
        contact_df = df[required_columns].copy()

        print(f"전체 데이터 수: {len(contact_df)}건")

        # 빈 값 처리 (NaN을 빈 문자열로 변환)
        contact_df = contact_df.fillna('')

        # 5개 컬럼 모두를 기준으로 중복 제거
        unique_contacts = contact_df.drop_duplicates(subset=required_columns, keep='first')

        print(f"중복 제거 후 데이터 수: {len(unique_contacts)}건")
        print(f"제거된 중복 데이터: {len(contact_df) - len(unique_contacts)}건")

        # 수요기관명으로 정렬 (가나다순)
        unique_contacts = unique_contacts.sort_values(by='수요기관명').reset_index(drop=True)

        # 출력 파일명 생성
        if output_file is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f'{timestamp}_담당자목록.xlsx'

        # 엑셀로 저장
        unique_contacts.to_excel(output_file, index=False, engine='openpyxl')
        print(f"\n새로운 엑셀 파일이 생성되었습니다: {output_file}")

        # 통계 정보 출력
        print("\n=== 통계 정보 ===")
        print(f"수요기관 수: {unique_contacts['수요기관명'].nunique()}개")
        print(f"담당자 이메일 수: {unique_contacts['공고기관담당자이메일주소'].apply(lambda x: x if x else None).nunique()}개")

        # 담당자 정보가 없는 데이터 확인
        empty_email = unique_contacts[unique_contacts['공고기관담당자이메일주소'] == '']
        if len(empty_email) > 0:
            print(f"\n이메일 정보가 없는 데이터: {len(empty_email)}건")

        empty_phone = unique_contacts[unique_contacts['공고기관담당자전화번호'] == '']
        if len(empty_phone) > 0:
            print(f"전화번호 정보가 없는 데이터: {len(empty_phone)}건")

        return unique_contacts

    except FileNotFoundError:
        print(f"오류: '{input_file}' 파일을 찾을 수 없습니다.")
    except Exception as e:
        print(f"오류 발생: {e}")


def analyze_contacts(df):
    """
    담당자 정보 분석 및 상세 출력

    Args:
        df: 담당자 정보 DataFrame
    """
    print("\n=== 상세 분석 ===")

    # 수요기관별 담당자 수
    print("\n[수요기관별 담당자 수]")
    org_count = df.groupby('수요기관명').size().sort_values(ascending=False)
    print(org_count.head(10))

    # 이메일이 있는 담당자 목록
    print("\n[이메일 목록 (중복 제거)]")
    emails = df[df['공고기관담당자이메일주소'] != '']['공고기관담당자이메일주소'].unique()
    print(f"총 {len(emails)}개의 이메일")
    for email in sorted(emails)[:20]:  # 상위 20개만 출력
        print(f"  - {email}")
    if len(emails) > 20:
        print(f"  ... 외 {len(emails) - 20}개")


def main():
    """메인 실행 함수"""

    # 입력 파일 경로
    input_file = '입찰공고(그래픽카드) - 20250101 ~ 20260131.xlsx'

    # 출력 파일명 (선택사항 - None이면 자동생성)
    output_file = '담당자_연락처_목록(그래픽카드 - 20250101 ~ 20260131).xlsx'

    print("=" * 60)
    print("입찰공고 담당자 정보 추출 및 중복 제거")
    print("=" * 60)

    # 담당자 목록 생성
    unique_contacts = create_unique_contact_list(input_file, output_file)

    # 상세 분석 (선택사항)
    if unique_contacts is not None:
        analyze_contacts(unique_contacts)

    print("\n작업이 완료되었습니다.")


if __name__ == "__main__":
    main()