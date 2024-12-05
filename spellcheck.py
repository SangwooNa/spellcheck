from hanspell import spell_checker

def check_spelling(text):
    """맞춤법 검사 함수"""
    result = spell_checker.check(text)
    if result.errors:
        print(f"Original: {text}")
        print(f"Corrected: {result.checked}")
        print(f"Errors: {result.errors}")
    else:
        print(f"No errors found in: {text}")

if __name__ == "__main__":
    # 검사할 텍스트를 입력하세요.
    sample_text = "이것은 맞춤법 검사 테스트입니다. 안돼 안돼!"
    check_spelling(sample_text)
