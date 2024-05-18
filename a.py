import matplotlib.pyplot as plt
import numpy as np
from matplotlib.gridspec import GridSpec
from matplotlib.colors import ListedColormap, BoundaryNorm

# Define the figure and GridSpec
fig = plt.figure(figsize=(10, 10))
gs = GridSpec(4, 1, figure=fig)  # Example with 4 rows

# Example data
data = np.random.randint(0, 3, size=(10, 10))
cmap = ListedColormap(['lightgrey', 'red', 'blue'])
norm = BoundaryNorm([-0.5, 0.5, 1.5, 2.5], cmap.N)

for i in range(4):
    ax = fig.add_subplot(gs[i, 0])
    cax = ax.matshow(data, cmap=cmap, norm=norm, aspect='auto')
    ax.spines['right'].set_visible(False)  # Hide right spine
    ax.spines['top'].set_visible(False)    # Hide top spine if desired
    ax.tick_params(axis='x', which='both', bottom=False, top=False, labelbottom=False)  # No x-axis ticks
    ax.tick_params(axis='y', which='both', left=False, right=False, labelleft=False)    # No y-axis ticks

# Comment out the colorbar addition
# plt.colorbar(cax, ax=fig.get_axes(), orientation='vertical')

fig.tight_layout()
plt.show()