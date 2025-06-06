import os
import torch
import pythoncom
from win32com.client import Dispatch, gencache

def get_kompas_api7():
    # module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    # api = module.IKompasAPIObject(
    #     Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
    #                                                              pythoncom.IID_IDispatch))
    # return module, api
    try:
        module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        api = module.IKompasAPIObject(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                     pythoncom.IID_IDispatch))
        return module, api
    except Exception as e:
        print(f"Error: {e}")
        return None, None

def get_kompas_api5():
    module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    api = module.KompasObject(            
        Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(module.KompasObject.CLSID,                                                        
                                                                 pythoncom.IID_IDispatch))   
    return module, api

def get_constants():
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    const_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants
    return const, const_3d

module7, api7 = get_kompas_api7()   # Подключаемся к API7 
module5, api5 = get_kompas_api5()   # Подключаемся к API5 
const, const_3d = get_constants()
app7 = api7.Application                     # Получаем основной интерфейс
app7.Visible = True                         # Показываем окно пользователю (если скрыто)
app7.HideMessage = const.ksHideMessageNo   # Отвечаем НЕТ на любые вопросы программы

#  Подключим константы API Компас
kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

#  Подключим описание интерфейсов API5
kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))

#  Подключим описание интерфейсов API7
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

Documents = application.Documents

import networkx as nx
import matplotlib.pyplot as plt

def model_to_graph(filePath):
    kompas_document = Documents.Open(filePath, True, False)
    if kompas_document is None:
        raise Exception(f"Не удалось открыть файл: {filePath}")
    # Инициализируем граф NetworkX
    G = nx.DiGraph()
    iDocument3D = api5.ActiveDocument3D()
    kompas_document_3d = module7.IKompasDocument3D(kompas_document)
    iPart7 = kompas_document_3d.TopPart
    iPart = iDocument3D.GetPart(const_3d.pTop_Part)


    faces = iPart.EntityCollection(0)

    for j in range (faces.GetCount()):
        entity = faces.GetByIndex(j)
        if entity != None:
            if entity.type == 6: # если поверхность
                entity.name = 'F_' + str(j)
            if entity.type == 7: # если ребро
                entity.name = 'E_' + str(j)
            if entity.type == 8: # если вершина
                entity.name = 'V_' + str(j)

    def get_entity(num):
        entity = faces.GetByIndex(num)
        if entity != None:
            if entity.type == 6: # если поверхность
                face = entity.GetDefinition()
                connected= face.ConnectedFaceCollection()
                # Добавляем грань как основной узел
                G.add_node(entity.name, type="face")
                for c in range(connected.GetCount()):
                    conFace = connected.GetByIndex (c)
                    if conFace.GetEntity().name == '':
                        print('!!')
                    if not G.has_edge(entity.name, conFace.GetEntity().name):  # Проверяем, нет ли уже такого ребра
                        G.add_edge(entity.name, conFace.GetEntity().name)

    for j in range (faces.GetCount()):
        get_entity(j)
    
    kompas_document.Close(False)
    return G

import torch_geometric
from torch_geometric.utils import from_networkx

def graph_to_data(G: nx.Graph, label: int = None) -> torch_geometric.data.Data:
    largest_scc = max(nx.strongly_connected_components(G), key=len)
    G = G.subgraph(largest_scc).copy()
    
    degrees = [val for _, val in G.degree()]
    closeness = list(nx.closeness_centrality(G).values())
    betweenness = list(nx.betweenness_centrality(G).values())
    pagerank = list(nx.pagerank(G).values())
    eccentricity = list(nx.eccentricity(G).values())
    
    features = torch.tensor(
        [[d, c, b, p, e] for d, c, b, p, e in zip(degrees, closeness, betweenness, pagerank, eccentricity)],
        dtype=torch.float
    )

    data = from_networkx(G)

    data.x = features
    if label is not None:
        data.y = torch.tensor([label], dtype=torch.long)

    return data

from pandas import read_csv
import time
dataset = []
df_path = os.path.join(os.getcwd(), "dataset.csv")
df = read_csv(df_path)
folders = os.listdir(os.path.join(os.getcwd(), "DSet"))

for f in folders:
    for idx, row in df.iterrows():
        try:
            filename = row["filename"]
            label = row["label"]
            modelPath = os.getcwd() + "\\DSet\\" + f + "\\" + filename
            if not os.path.exists(modelPath):
                print(f"Файл не существует: {modelPath}")
                continue
            G = model_to_graph(modelPath)
            dataset.append(graph_to_data(G, label))
        except Exception as ex:
            print(f"Error: {ex}")
            break
        else:
            print(f"got it! - {modelPath, label}")
        time.sleep(0.3)
            
from torch_geometric.loader import DataLoader

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
    
model = GNN(in_channels=5, hidden_channels=64, num_classes=4)
optimizer = torch.optim.Adam(model.parameters(), lr=0.01)
criterion = torch.nn.CrossEntropyLoss()

for epoch in range(1, 900):
    model.train()
    total_loss = 0

    for data in loader: 
        optimizer.zero_grad()
        out = model(data.x, data.edge_index, data.batch)
        loss = criterion(out, data.y)
        loss.backward()
        optimizer.step()
        total_loss += loss.item()

    if epoch % 100 == 0:
        print(f"Epoch {epoch}, Loss: {total_loss:.4f}")