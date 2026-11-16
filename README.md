# 현장 체크 지도

담당 건물을 지도에 띄우고, 층별 임차인이 그대로인지/바뀌었는지 폰으로 체크하는 사이트.

- 사이트: https://bhsbse-create.github.io/field-check-map/
- 데이터는 비밀번호로 암호화된 `data.enc.json` 만 올라가 있음. **엑셀 원본·평문 데이터는 절대 커밋하지 않기** (`.gitignore`가 막아 둠)
- 체크 기록은 각 폰(브라우저)에만 저장됨 → 메뉴(⋯)에서 CSV로 받아서 회사 엑셀에 옮기기

## 파일
| 파일 | 내용 |
|---|---|
| `index.html` | 사이트 화면·기능 전부 (지도, 건물 화면, 목록, 체크, CSV 내보내기) |
| `data.enc.json` | 암호화된 건물·층 데이터 |
| `tools/update.py` | 새 엑셀로 데이터 다시 만드는 스크립트 |

## 새 엑셀로 데이터 갱신하기
엑셀 파일은 **이 폴더 밖**(예: 바탕화면)에 두고:

```
pip install openpyxl cryptography
python tools/update.py --excel "C:/Users/내이름/Desktop/2026_.xlsx" --password 지금비번
git add data.enc.json
git commit -m "데이터 갱신"
git push
```

- 몇 분 뒤 사이트에 반영되고, 폰에서는 비밀번호를 한 번 다시 입력해야 함
- 기존 건물 좌표는 그대로 재사용, 새 건물만 좌표를 찾음
- 비밀번호 바꾸기: `--password 새비번 --old-password 예전비번`
- 위치를 전부 정확하게: 카카오 REST 키를 받아서 `--kakao-key 키 --regeocode`
- 담당자가 바뀌면 `--name 이름`

## 클로드 코드에게 시킬 때
"tools/update.py 로 바탕화면의 새 엑셀을 반영해서 푸시해 줘. 비번은 ○○○○" 처럼 말하면 됨.
화면을 고칠 땐 `index.html` 만 수정하면 됨.
