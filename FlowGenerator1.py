import pandas as pd
import graphviz
from graphviz import Digraph

# Step 1: Read Data from Excel
df = pd.read_excel('FinalSourceDestination.xlsx')  # Update with your file path

# Step 2: Define Nodes and Edges
nodes = set(df[['Domain', 'Destination']].stack().astype(str))
edges = set((str(df.loc[i, 'Domain']), str(df.loc[i, 'Destination'])) for i in range(len(df)))

# Step 3: Preprocess Nodes to Remove Duplicates
processed_nodes = set()
unique_nodes = []
for node in nodes:
    if node not in processed_nodes:
        processed_nodes.add(node)
        unique_nodes.append(node)

# Step 4: Create Flowchart
dot = Digraph(engine='dot')
dot.attr(size='100,100')
dot.attr('node', shape='rectangle', style='filled', fillcolor='lightblue', fontname='Arial', fontsize='10')
dot.attr('edge', color='gray', arrowhead='vee', penwidth='1')

# Group similar nodes
node_groups = {}
for node in unique_nodes:
    group_key = node.split('_')[0]  # Assuming the first part of the node name represents the group
    node_groups.setdefault(group_key, []).append(node)

# Organize nodes into rows
rows = list(node_groups.values())

# Add nodes to the graph, organizing them into rows
for i, row_nodes in enumerate(rows):
    with dot.subgraph() as s:
        s.attr(rank='same')
        for node in row_nodes:
            s.node(node, label=node, fontsize='10', shape='box')

# Add edges
for edge in edges:
    dot.edge(*edge)

# Step 5: Render and Save
dot.render('output1', format='png', cleanup=True)