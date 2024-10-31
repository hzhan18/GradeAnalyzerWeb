import matplotlib
matplotlib.use('Agg')  # Use a non-interactive backend suitable for Flask
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import os

# Set font properties to support Chinese characters
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # Adjust as needed
plt.rcParams['axes.unicode_minus'] = False  # Ensure minus signs display correctly

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
    # Check if distribution has the correct structure
    if not isinstance(distribution, dict) or not all(isinstance(d, dict) and '人数' in d for d in distribution.values()):
        raise ValueError("Invalid distribution format: Expected a dictionary with '人数' key in each value.")

    scores = list(distribution.keys())
    counts = [d['人数'] for d in distribution.values()]
    
    # Create the plot
    plt.figure(figsize=(10, 6))
    plt.bar(scores, counts, color='skyblue')
    plt.xlabel('分数段')
    plt.ylabel('人数')
    plt.title(title)
    plt.gca().yaxis.set_major_locator(MaxNLocator(integer=True))
    plt.grid(True)
    
    # Ensure the directory exists
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    
    # Save the plot to the specified file path
    plt.savefig(file_path)
    plt.close()
    
    # Return the file path for web access
    return file_path
