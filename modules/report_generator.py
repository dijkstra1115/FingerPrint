from modules import original_template

class Finger:
    # 設定紋型排名映射
    pattern_mapping = {
        "Wp": 11, "We": 10, "Wt": 9, "Ws": 8, "Wc": 7,
        "U": 6, "R": 5, "Au": 4, "Ar": 3, "As": 2, "At": 1
    }

    def __init__(self, name, pattern, left_value, right_value):
        # 或許可以將 self.name 拆成 self.hand 以及 self.number
        # 目前只有在 self.calculate_hand_weight() 會使用到 self.name 這個屬性
        self.name = name
        self.pattern = pattern
        self.left_value = left_value
        self.right_value = right_value
        self.value = max(left_value, right_value)
        self.percentage = 0.0
        self.pattern_weight = self.pattern_mapping.get(pattern, float('inf'))

        # 相同紋型在不同手上的權重
        self.hand_weight = self.calculate_hand_weight()
        # 相同紋型在同一手上的排名
        self.hand_ranking = 0

        self.ranking = 0
        self.tmp_ranking = 0
        self.peer = None  # 初始化 peer 為 None

    @property
    def peer_ranking(self):
        # 如果 peer 存在，則返回 peer 的 rank；否則返回 None
        if self.peer:
            return self.peer.ranking
        else:
            return None

    @property
    def peer_tmp_ranking(self):
        if self.peer:
            return self.peer.tmp_ranking
        else:
            return None
    
    @peer_ranking.setter
    def peer_ranking(self, value):
        # 如果 peer 存在，則同時設定 peer 的 rank
        if self.peer:
            self.peer.ranking = value

    @peer_tmp_ranking.setter
    def peer_tmp_ranking(self, value):
        if self.peer:
            self.peer.tmp_ranking = value

    def set_peer(self, peer_finger):
        # 設定 peer 指向對應的手指
        self.peer = peer_finger
        peer_finger.peer = self

    def calculate_hand_weight(self):
        R_priority = ["Wp", "We", "Wt", "Ws", "Wc", "As", "At"]
        L_priority = ["U", "R", "Au", "Ar"]
        if self.name[:1] == 'R' and self.pattern in R_priority:
            return 2
        elif self.name[:1] == 'L' and self.pattern in L_priority:
            return 2
        else:
            return 1

def generate_report(user_name, pricing_plan, data):

    pattern_map = {
        "We": "Ws",
        "Wt": "Ws",
        "Ws": "Ws",
        "Wp": "Wc",
        "Wc": "Wc",
        "U" : "U",
        "R" : "R",
        "As": "A",
        "Au": "A",
        "Ar": "A",
        "At": "A",
    }

    fingers = {}
    for key in data.keys():
        # 只处理每组手指数据的代码部分
        if '_code' in key:
            finger_name = key.split('_')[0]  # 如 "R1", "R2" 等
            # 创建 Finger 实例
            fingers[finger_name] = Finger(
                finger_name,
                data[f'{finger_name}_code'],
                int(data[f'{finger_name}_left_value']),
                int(data[f'{finger_name}_right_value'])
            )

    original_template.customize_report(user_name, pricing_plan, fingers, pattern_map)