import pandas as pd
import matplotlib.pyplot as plt

def draw_plots(file, plot_title,save_title):
    df = pd.read_csv(file)
    fig, ax = plt.subplots()
    index = df['scenario']
    bar_width = 0.35
    opacity = 0.8

    rects1 = plt.bar(index, df['hcm'], bar_width,
    alpha=opacity,
    color='red',edgecolor='black', hatch="/",
    label='HCM')

    rects2 = plt.bar(index + bar_width, df['proposed'], bar_width,
    alpha=opacity,
    color='g',edgecolor='black', hatch="\\",
    label='Proposed Model')

    plt.xlabel('Scenario')
    plt.ylabel('Cycle Length (s)')
    plt.title(plot_title)
    plt.xticks(index + bar_width, ('1','2','3','4','5','6','7','8','9'))
    plt.legend()

    plt.tight_layout()
    plt.show()
    fig.savefig(save_title, dpi=600)

file = '../for_viz_1_cycles.csv'
plot_title = 'First Intersection: Cycle Length Comparisons'
save_title = '../first_cycles.pdf'
draw_plots(file,plot_title,save_title)

file = '../for_viz_2_cycles.csv'
plot_title = 'Second Intersection: Cycle Length Comparisons'
save_title = '../second_cycles.pdf'
draw_plots(file,plot_title,save_title)


file = '../for_viz_3_cycles.csv'
plot_title = 'Third Intersection: Cycle Length Comparisons'
save_title = '../third_cycles.pdf'
draw_plots(file,plot_title,save_title)
