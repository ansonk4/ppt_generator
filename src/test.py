from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_AXIS_CROSSES
from pptx.util import Inches
from pptx.enum.chart import XL_LEGEND_POSITION

# Create a new presentation
prs = Presentation()

# Add a slide with a blank layout
slide_layout = prs.slide_layouts[6]  # Blank layout
slide = prs.slides.add_slide(slide_layout)

# Sample data for the chart
categories = ['Q1', 'Q2', 'Q3', 'Q4']
values = [20, 35, 15, 40]

# Create chart data
chart_data = CategoryChartData()
chart_data.categories = categories
chart_data.add_series('Sales', values)

# Add chart to slide
x, y, cx, cy = Inches(2), Inches(1.5), Inches(6), Inches(4.5)
chart = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
).chart

# Configure the chart to make bars hang from top
# The most effective approach is to use negative values and flip the axis
value_axis = chart.value_axis
category_axis = chart.category_axis

# Set the value axis to cross at maximum (top of chart)
# category_axis.crosses = XL_AXIS_CROSSES.MAXIMUM

# Optional: Reverse the value axis so values increase downward
value_axis.reverse_order = True

# Set appropriate scale to ensure bars start from top
max_value = max(values)
value_axis.minimum_scale = 0
value_axis.maximum_scale = max_value * 1.2  # Add some padding

# Configure chart appearance
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.RIGHT

# Set chart title
chart.chart_title.text_frame.text = "Hanging Bar Chart"

# Optional: Format the bars
plot = chart.plots[0]
plot.has_data_labels = True

# Access the first series to customize bar appearance
series = plot.series[0]
# You can customize colors, etc. here if needed

# Save the presentation
prs.save('hanging_bar_chart.pptx')
print("Chart created successfully!")

# Alternative approach: Using negative values to create hanging effect
print("\n--- Alternative Method with Negative Values ---")

# # Create another slide with negative values approach
# slide2 = prs.slides.add_slide(slide_layout)

# # Convert positive values to negative to create hanging effect
# hanging_values = [-v for v in values]

# # Create new chart data with negative values
# chart_data2 = CategoryChartData()
# chart_data2.categories = categories
# chart_data2.add_series('Sales', hanging_values)

# # Add second chart
# chart2 = slide2.shapes.add_chart(
#     XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data2
# ).chart

# # Configure the second chart
# value_axis2 = chart2.value_axis
# value_axis2.minimum_scale = min(hanging_values) * 1.1
# value_axis2.maximum_scale = 0

# # Position category axis at top for the negative values method
# category_axis2 = chart2.category_axis
# # category_axis2.crosses = XL_AXIS_CROSSES.MAXIMUM

# # Set title
# chart2.chart_title.text_frame.text = "Hanging Bars (Negative Values Method)"
# chart2.has_legend = True

# # Add data labels
# chart2.plots[0].has_data_labels = True

# # Save with both approaches
# prs.save('hanging_bar_charts_both_methods.pptx')
# print("Both chart methods created successfully!")