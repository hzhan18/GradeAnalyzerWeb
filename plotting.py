import matplotlib
matplotlib.use('Agg')  # Use a non-interactive backend suitable for Flask
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import os
from matplotlib.font_manager import FontProperties

# 设置中文字体的路径
font_path = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
my_font = FontProperties(fname=font_path)  # 加载指定路径的字体

def plot_distribution(distribution, title, file_path):
    """
    Plots a distribution of scores and saves it to the specified file path.

    Parameters:
        distribution (dict): Dictionary where keys are score ranges (e.g., '60-70') and
                             values are dictionaries with '人数' (count of people) and 
                             possibly other statistics.
        title (str): Title of the plot.
        file_path (str): File path to save the plot image.
    """
    try:
        # Check if distribution has the correct structure
        if not isinstance(distribution, dict) or not all(isinstance(d, dict) and '人数' in d for d in distribution.values()):
            raise ValueError("Invalid distribution format: Expected a dictionary with '人数' key in each value.")

        scores = list(distribution.keys())
        counts = [d['人数'] for d in distribution.values()]
        
        # Create the plot
        plt.figure(figsize=(10, 6))
        plt.bar(scores, counts, color='skyblue')
        plt.xlabel("分数段", fontproperties=my_font)  # 设置X轴字体
        plt.ylabel("人数", fontproperties=my_font)    # 设置Y轴字体
        plt.title(title, fontproperties=my_font)     # 设置标题字体
        plt.xticks(fontproperties=my_font)           # 设置X轴标签字体，确保分数段使用中文字体
        plt.gca().yaxis.set_major_locator(MaxNLocator(integer=True))
        plt.grid(True)
        
        # Ensure the directory exists
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        # Save the plot to the specified file path
        plt.savefig(file_path)
        plt.close()
        
        # Return the file path for web access
        return file_path

    except Exception as e:
        print(f"An error occurred while generating the plot: {e}")
        raise
