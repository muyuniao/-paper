import pandas as pd

# 定义参数
m = 0.502
C = 1.0
epsilon = 0.90
g = 9.81
B = 12
sigma_s = 1.0

# 创建空列表用于存储 H_W 和 Q 的值
H_W_values = []
Q_values = []

# 计算 H_W 从 0 到 19 的 Q 值
for H_W in range(20):
    H_W_values.append(H_W)
    Q = C * m * epsilon * sigma_s * B * (2 * g) ** 0.5 * H_W ** (3 / 2)
    Q_values.append(Q)

# 创建 DataFrame
data = {'H_W': H_W_values, 'Q': Q_values}
df = pd.DataFrame(data)

# 保存为 Excel 文件
df.to_excel('flow_calculation.xlsx', index=False)