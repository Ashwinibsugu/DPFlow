import pandas as pd
from graphviz import Digraph

# Step 1: Read Data from Excel
df = pd.read_excel('FinalSourceDestination.xlsx')  # Update with your file path

# Step 2: Define Nodes and Edges
nodes = set(df[['Domain', 'Destination']].stack().astype(str))
edges = {(str(df.loc[i, 'Domain']), str(df.loc[i, 'Destination'])) for i in range(len(df))}

# Step 3: Preprocess Nodes to Group Nodes with Similar Names
node_groups = {}
for node in nodes:
    prefix = node.split('_')[0]  # Extract the prefix of the node name
    node_groups.setdefault(prefix, set()).add(node)

# Step 4: Create Flowchart
dot = Digraph(engine='dot')
dot.attr(size='100,100')
dot.attr('node', shape='rectangle', style='filled', fillcolor='lightblue', fontname='Arial', fontsize='10')
dot.attr('edge', color='gray', arrowhead='vee', penwidth='1')

# Step 5: Add nodes to the graph, ensuring each node name appears only once
for group_prefix, group_nodes in node_groups.items():
    with dot.subgraph() as s:
        s.attr(rank='same')
        for node in group_nodes:
            s.node(node, label=node, fontsize='10', shape='box')

# Step 6: Add edges
for edge in edges:
    dot.edge(*edge)

# Step 7: Render and Save
dot.render('output', format='png', cleanup=True)