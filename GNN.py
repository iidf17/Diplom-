from torch_geometric.loader import DataLoader
from k_parser
from pandas import read_csv
dataset = []
df_path = os.getcwd() + "\\filesK\\dataset.csv"
df = read_csv(df_path)
for idx, row in df.iterrows():
    filename = row["filename"]
    label = row["label"]
    filepath = os.path.join(os.getcwd() + "\\filesK", filename)
    G = model_to_graph(filepath)
    dataset.append(graph_to_data(G, label))

loader = DataLoader(dataset, batch_size=8, shuffle=True)

import torch
import torch.nn.functional as F
from torch_geometric.nn import GCNConv, global_mean_pool

class GNN(torch.nn.Module):
    def __init__(self, in_channels, hidden_channels, num_classes):
        super(GNN, self).__init__()
        self.conv1 = GCNConv(in_channels, hidden_channels)
        self.conv2 = GCNConv(hidden_channels, hidden_channels)
        self.lin = torch.nn.Linear(hidden_channels, num_classes)

    def forward(self, x, edge_index, batch):
        # x — признаки узлов
        # edge_index — связи
        # batch — какой узел к какому графу относится

        x = self.conv1(x, edge_index)
        x = F.relu(x)
        x = self.conv2(x, edge_index)

        # Агрегация признаков всех узлов в один вектор на граф
        x = global_mean_pool(x, batch)

        x = self.lin(x)
        return x
    

model = GNN(in_channels=5, hidden_channels=64, num_classes=2)
optimizer = torch.optim.Adam(model.parameters(), lr=0.01)
criterion = torch.nn.CrossEntropyLoss()

for epoch in range(1, 300):
    model.train()
    total_loss = 0

    for data in loader:  # один батч (несколько графов)
        optimizer.zero_grad()
        out = model(data.x, data.edge_index, data.batch)  # data.batch нужен для глобального пула
        loss = criterion(out, data.y)
        loss.backward()
        optimizer.step()
        total_loss += loss.item()

    if epoch % 40 == 0:
        print(f"Epoch {epoch}, Loss: {total_loss:.4f}")

testFilePath = os.getcwd() + "\\filesK\\cube_test.m3d"
data = model_to_graph(testFilePath)
test_graph = graph_to_data(data)

with torch.no_grad():
    out = model(test_graph.x, test_graph.edge_index, test_graph.batch)
    pred_class = out.mean(dim=0).argmax().item()  # среднее по всем узлам
    print(f"Предсказанный класс: {pred_class}")
    