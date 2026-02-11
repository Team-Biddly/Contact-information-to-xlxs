import requests
import pandas as pd


def get_pps_rental_data():
    service_key = ""
    base_url = "http://apis.data.go.kr/1230000/ao/OrderPlanSttusService"
    search_term = ""

    operations = [
        "getOrderPlanSttusListThngPPSSrch",
        "getOrderPlanSttusListCnstwkPPSSrch",
        "getOrderPlanSttusListServcPPSSrch",
        "getOrderPlanSttusListFrgcptPPSSrch"
    ]

    # 異붿텧�� 而щ읆 留ㅽ븨 (�곷Ц紐�: �쒓�紐�)
    column_mapping = {
        'bsnsDivNm': '',
        'orderInsttCd': '',
        'totlmngInsttNm': '',
        'jrsdctnDivNm': '',
        'orderInsttNm': '',
        'prcrmntMethd': '',
        'bizNm': '',
        'cntrctMthdNm': '',
        'sumOrderAmt': '',
        'deptNm': '',
        'ofclNm': '',
        'telNo': '',
        'nticeDt': '',
        'orderPlanUntyNo': ''
    }

    all_items = []

    for op in operations:
        print(f"{op} ...")

        params = {
            'ServiceKey': service_key,
            'bizNm': search_term,
            'numOfRows': '100',
            'pageNo': '1',
            'type': 'json',
            'orderBgnYm': '202509',
            'orderEndYm': '202612',
            'inqryBgnDt': '202509010000',
            'inqryEndDt': '202612312359'
        }

        try:
            # API �몄텧
            response = requests.get(f"{base_url}/{op}", params=params)

            if response.status_code == 200:
                data = response.json()
                body = data.get('response', {}).get('body', {})
                items = body.get('items', [])

                if isinstance(items, list) and items:
                    all_items.extend(items)
                elif isinstance(items, dict) and 'item' in items:
                    item_data = items['item']
                    if isinstance(item_data, list):
                        all_items.extend(item_data)
                    else:
                        all_items.append(item_data)
            else:
                print(f"{op} (HTTP {response.status_code})")
        except Exception as e:
            print(f"{op} : {e}")

    if all_items:
        df = pd.DataFrame(all_items)

        available_cols = [col for col in column_mapping.keys() if col in df.columns]
        final_df = df[available_cols].rename(columns=column_mapping)

        output_filename = f"{search_term}_.xlsx"
        final_df.to_excel(output_filename, index=False)
        print(f"\n{len(final_df)} '{output_filename}'.")
    else:
        print("\n")


if __name__ == "__main__":
    get_pps_rental_data()