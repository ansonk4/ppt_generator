from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LABEL_POSITION
from pptx.util import Inches

def apply_data_labels(chart):
        plot = chart.plots[0]
        plot.has_data_labels = True
        for series in plot.series:
            values = series.values
            counter = 0
            for point in series.points:
                data_label = point.data_label
                data_label.has_text_frame = True
                data_label.text_frame.text = str(values[counter])
                counter = counter + 1

def create_donut_chart_ppt():
    # Create a new presentation
    prs = Presentation()
    
    # Add a slide with a blank layout
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add a title to the slide
    title_shape = slide.shapes.add_textbox(
        left=Inches(1), 
        top=Inches(0.5), 
        width=Inches(8), 
        height=Inches(1)
    )
    title_frame = title_shape.text_frame
    title_frame.text = "Sales by Region - Donut Chart"
    
    # Sample data for the donut chart
    chart_data = CategoryChartData()
    chart_data.categories = ['North', 'South', 'East', 'West', 'Central']
    chart_data.add_series('Sales', (35, 25, 20, 15, 5))
    # Add chart to slide
    x, y, cx, cy = Inches(1.5), Inches(2), Inches(7), Inches(5)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT, x, y, cx, cy, chart_data
    ).chart
    
    # Customize the chart
    chart.has_legend = True
    
    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.show_category_name = True
    data_labels.show_percentage = True
    data_labels.show_value = False

    # for series in plot.series:
    #     values = series.values
    #     counter = 0
    #     for point in series.points:
    #         data_label = point.data_label
    #         data_label.has_text_frame = True
    #         data_label.text_frame.text = str(values[counter])
    #         counter = counter + 1

    # Save the presentation
    prs.save('donut_chart_presentation.pptx')
    print("PowerPoint presentation with donut chart created successfully!")
    print("File saved as: donut_chart_presentation.pptx")

if __name__ == "__main__":
    create_donut_chart_ppt()