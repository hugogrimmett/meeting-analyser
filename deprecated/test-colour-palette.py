import matplotlib.pyplot as plt
import colorsys
import math

def distinct_color_grid(n):
    """
    Generate n visually distinct RGB colors by using a grid in HLS space.
    """
    k = int(math.ceil(n ** 0.5))  # number of hues
    l_values = [0.45, 0.65, 0.8] if n > 20 else [0.5, 0.7]
    colors = []
    for l in l_values:
        for i in range(k):
            h = i / float(k)
            s = 0.7
            rgb = colorsys.hls_to_rgb(h, l, s)
            colors.append(rgb)
            if len(colors) >= n:
                break
        if len(colors) >= n:
            break
    return colors[:n]

def make_global_color_dict(participants):
    n = len(participants)
    colors = distinct_color_grid(n)
    return dict(zip(participants, colors))

# Test and visualize for 30 colors
palette = distinct_color_grid(30)
for i, color in enumerate(palette):
    plt.plot([0,1], [i,i], color=color, linewidth=15)
plt.yticks([])
plt.show()
