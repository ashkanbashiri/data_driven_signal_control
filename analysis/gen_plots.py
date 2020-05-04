import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
'''
sa_results = []
sa_results = [[0.03387744, 0.04218781, 0.03423842, 0.04578994, 0.26688983],
         [0.33675454, 0.31773746, 0.36306585, 0.44983512, 0.69078965],
         [0.04643806, 0.00493544, 0.07501189, 0.01683896, 0.22284424],
         [0.38283312, 0.39357963, 0.49206298, 0.42272254, 0.52532446],
         [0.0381721,  0.04096238, 0.04734618, 0.04511934, 0.25572823],
         [0.35267031, 0.32047399, 0.37047719, 0.43297962, 0.66397433],
         [0.04686533, 0.00715493, 0.08123541, 0.01625125, 0.22802752],
         [0.3746252,  0.37882559, 0.49343411, 0.423958,   0.53557161]]
sa_results_df = pd.DataFrame(sa_results, columns = ['flow_1', 'flow_2','flow_3', 'flow_4', 'lost_time'],
                             index=['ph1_first_order','ph1_total_order',
                                    'ph2_first_order','ph2_total_order',
                                    'ph3_first_order','ph3_total_order',
                                    'ph4_first_order','ph4_total_order'])
sa_results_df['index_order'] = ['first_order', 'total_order',
                                'first_order', 'total_order',
                                'first_order', 'total_order',
                                'first_order', 'total_order']
'''
sa_results = pd.read_csv('../sim_results/sa_results.csv')
sa_results['pred_index'] = sa_results['predictor'].astype(str) + '_' + sa_results['index_order'].astype(str)
sa_results['sens']= sa_results['pred_index'].map(sa_results.groupby(['pred_index'])['sensitivity'].mean())
sa_results['predictor_phase'] = sa_results['predictor'].astype(str) + '__PhaseGroup' + sa_results['phase_group'].astype(str)
print(sa_results.head())
f = plt.figure(figsize=(100, 100))

sns.set_color_codes("pastel")
g = sns.catplot(y="predictor_phase", x="sensitivity", hue="index_order", data=sa_results,
                height=6, kind="bar", palette="muted", legend=False, orient='h')
g.despine(left=True)
g.set_ylabels("Sensitivity")
plt.legend(loc='upper right')
plt.title('Sensitivity Analysis')
plt.show()
g.savefig("../sa_results.pdf", bbox_inches='tight')
