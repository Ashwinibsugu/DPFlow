import pandas as pd
from graphviz import Digraph

# Step 1: Read Data from Excel
try:
    df = pd.read_excel('FinalSourceDestination.xlsx')  # Update with your file path
except FileNotFoundError:
    print("Error: Excel file not found.")
    exit()

# Step 2: Extract Unique Nodes and Edges
nodes = set()

# Extract unique nodes from 'Domain' and 'Destination' columns
for column in ['Domain', 'Destination']:
    nodes.update(df[column].astype(str).str.strip())  # Remove leading and trailing whitespace

# Step 3: Group Nodes Based on Prefix
node_groups = {}
for node in nodes:
    group_key = node.split('_')[0]  # Extract the prefix
    node_groups.setdefault(group_key, set()).add(node)

# Step 4: Create Flowchart
dot = Digraph(engine='dot')
dot.attr(size='100,100')
dot.attr('node', shape='rectangle', style='filled', fillcolor='lightblue', fontname='Arial', fontsize='10')
dot.attr('edge', color='gray', arrowhead='vee', penwidth='1')

# Add nodes to the graph, organizing them into subgraphs based on prefix
for group_key, group_nodes in node_groups.items():
    with dot.subgraph() as s:
        s.attr(rank='same')  # Ensure nodes in the same subgraph are on the same rank
        s.attr(label=group_key)  # Set subgraph label to the prefix
        for node in group_nodes:
            s.node(node, label=node, fontsize='10', shape='box')

# Add edges to the graph
added_edges = set()  # Track added edges to avoid duplicates
for _, row in df.iterrows():
    source = str(row['Domain']).strip()
    destination = str(row['Destination']).strip()
    if (destination, source) not in added_edges:  # Check if the reverse edge already exists
        dot.edge(source, destination)  # Remove leading and trailing whitespace
        added_edges.add((source, destination))

# Step 5: Render and Save
dot.render('output1', format='png', cleanup=True)
