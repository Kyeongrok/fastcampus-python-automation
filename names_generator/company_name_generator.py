from random import randrange, sample

def generate_company_name(n=100):
    result = set()
    prefixs = list({'파인', '이클래스', '심온', '다온', '탑', '화이트', '그레이스', '빈센트', '원진', '오케이', '다인', '프라임', '을지', '에이스', '그랑', '큐원', '라임', '바닐라', '에버엠', '아이비', '다솜', '루덴', '밀레니엄', '아너스', '퍼스트', '봄', '예담', '메이플', '보스톤', '클리오', '벨', '예담', '지앤미', '단아',
        '세실', '메트로', '더블유', '신도', '신도', '아이리스', '휴먼', '세명'})
    postfixs = list({'파트너스', '홀딩스', '플레이어', '컴퍼니', '코리아', '리테일', '컨세션', '마켓',
                '마트', '텍', '테크', '컨설팅', '메디', '메디컬', '캐미컬'})
    # print(postfixs)
    cnt = 0
    while len(result) < n:
        prefix = prefixs[randrange(len(prefixs))]
        postfix = postfixs[randrange(len(postfixs))]
        result.add(f'{prefix}{postfix}')
        cnt += 1
    print('try:', cnt)
    return result


if __name__ == '__main__':
    company_names = generate_company_name()
    print(company_names)
