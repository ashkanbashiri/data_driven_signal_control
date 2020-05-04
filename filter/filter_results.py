import pandas as pd


def filter_best(df, to_predict, predictors, metric):
    keep_list = to_predict + predictors + [metric] + ['totalflow(Veh/h)']
    df = df.sort_values(metric, ascending=True).drop_duplicates(predictors).reset_index()

    drop_list = list(set(df.columns)-set(keep_list))
    #df = df.drop(drop_list, axis=1)
    return df

