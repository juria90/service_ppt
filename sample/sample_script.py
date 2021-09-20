# sample script to process input(list of string) and return output(list of string)
# 

sample_text = """1. 교회 창립 멤버들...

2. 아둘람 동굴에 모여든 사람들
* 동굴로 피신한 기진맥진한 다윗
* 그에게 찾아 온 400 여 명의 사람들...
* 다윗 왕국을 세워가는 중추적 인물들

3. 그들이 하나님 나라의 용사가 될 수 있었던 이유
* 하나님의 은혜를 알았기 때문에
* 서로 사랑으로 하나 되었기 때문에
* 침침한 동굴에서도 하나님의 영광을 보았기 때문에

4. 알고 보면 모든 것이 하나님의 섭리이다
* 지난 시간 내 인생에 대한 해석
* 미가 선지자의 예언
"""

input = sample_text.split("\n\n")

script = r"""output = []
for sec in input:
    lns = sec.splitlines()
    lns0 = lns[0]
    lns1_ = lns[1:]
    output.append(lns0)
    if len(lns1_):
        output.extend([lns0 + "\n" + l for l in lns1_])
"""

def run_script(script, input):
    try:
        gdict = {"input": input}
        exec(script, gdict)
        return gdict["output"]
    except Exception as e:
        print("Error: %s" % e)
        pass

output = run_script(script, input)
print(output)
