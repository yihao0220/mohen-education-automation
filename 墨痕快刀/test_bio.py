import sys
sys.path.append('E:/PYTHON/practice/墨痕教育/core')
import config

test_text = """1.细胞代谢
1．(14分)(2024·岳阳高三模拟)科研人员发现植物的细胞呼吸除具有与动物细胞相同的途径外，还包含另一条借助交替氧化酶(AOX)的途径。请回答下列问题：
(1)交替氧化酶(AOX)分布在植物细胞线粒体内膜上
(2)进一步研究表明，AOX途径可能与光合作用有关
2．(16分)为探究冰叶日中花在干旱和盐胁迫下的响应和适应能力
(1)据表可知，中度干旱胁迫对冰叶日中花的影响体现在
(2)研究表明，在盐胁迫下，冰叶日中花光合作用的暗反应过程从C3途径
(3)该实验需在培养过程中保持这四组实验的光照
3．(14分)(2024·张家界高三模拟)土壤盐分过高对植物的伤害作用称为盐胁迫
(1)与第3天相比，第6天盐胁迫组竹柳苗的净光合速率下降
(2)重度盐胁迫条件下，12天内竹柳苗干重的变化为
(3)研究发现，盐胁迫下竹柳根细胞内的脯氨酸含量明显升高
(4)进一步研究发现，盐胁迫下竹柳根细胞内的脯氨酸含量增加
4．(16分)光合作用机理是作物高产的重要理论基础
(1)测得两种水稻分别在弱光照和强光照条件下净光合速率的变化如图1
(2)通常情况下，叶绿素含量与植物的光合速率呈正相关"""

lines = test_text.strip().split('\n')
scores = {name: 0 for name in config.CONFIG_SCIENCE["formats"]}

import re

for line in lines:
    line = line.strip()
    if not line: continue
    
    is_obs = False
    for pat in config.CONFIG_SCIENCE['obstacles']:
        if re.match(pat, line):
            is_obs = True
            break
            
    if is_obs: continue
    
    for name, pattern in config.CONFIG_SCIENCE["formats"].items():
        if re.match(pattern, line):
            scores[name] += 1
            
print(scores)

from core_parser import detect_best_format
class MockRange:
    def __init__(self, t): self.Text = t
    def Information(self, *args): return False

class MockPara:
    def __init__(self, t): self.Range = MockRange(t)

class MockParas:
    def __init__(self, lines):
        self.paras = [MockPara(l) for l in lines]
    def __call__(self, idx):
        return self.paras[idx-1]

class MockDoc:
    def __init__(self, lines):
        self.Paragraphs = MockParas(lines)

doc = MockDoc(lines)
config.CURRENT_CONFIG = config.CONFIG_SCIENCE
fmt, name = detect_best_format(doc, 1, len(lines))
print(f"Detected format: {name} (pattern: {fmt})")

