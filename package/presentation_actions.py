import logging

def scan_for_shapes(slide, debug=False):
    placeholder_count = 0

    for shape in slide.Shapes:
        if hasattr(shape, 'PlaceholderFormat'):
            placeholder_count += 1
            logging.debug(f"scan_for_shapes found a Placeholder: ID {shape.Id}, Name: {shape.Name}")  

    logging.debug(f"scan_for_shapes found a total of: {placeholder_count} Placeholders")

    return placeholder_count

def skip_to_page(presentation, slide_number):

    num_slides = presentation.Slides.Count
    assert num_slides >= slide_number, f"you are trying to skip to slide {slide_number}, the highest page number is {num_slides}"

    try:
        slide_show_view = presentation.SlideShowWindow.View
        #slide_show_view.Next()
        slide_show_view.GotoSlide(slide_number)
    except AttributeError:
        logging.critical("No active slideshow found.")


def collect_group_shapes(slide):
    group_shape_list = []
    
    for shape in slide.Shapes:
        if "Group" in shape.Name:
            group_shape_list.insert(0, shape)  #win32 detects from background layer to the front. we want the header at index 0
            logging.debug(f"collect_group_shapes found a Group: ID {shape.Id}, Name: {shape.Name}")
    
    return group_shape_list


def populate_group(group, contents):

    assert "Group" in group.Name, "you are trying to add text to a non group object"

    content_index = 0
    for placeholder_index in range(1, group.GroupItems.Count+1):  # PowerPoint collections are 1-indexed
        if content_index >= len(contents):
            logging.error("there are more group items than content")
            break
        content_placeholder = group.GroupItems.Item(placeholder_index)
        if not "TextBox" in content_placeholder.Name:
            logging.debug("skipped adding content to a rectangle")
            content_index -= 1
            pass

        content = contents[content_index]
        content_placeholder.TextFrame.TextRange.Text = content
        content_index += 1
    
    assert content_index <= len(contents), "more content than placeholders in populate_group()"

def add_content_to_group_shapes(group_shape_list, content_per_column):
    for group_shape in group_shape_list:

        assert len(content_per_column) == group_shape.GroupItems.Count, "content and group must have the same amount of elements"

        for i, content_placeholder in enumerate(group_shape.GroupItems):
            if "TextBox" in content_placeholder.Name:
                content_placeholder.TextFrame.TextRange.Text = content_per_column[i]
            # Additional logic can be added here to ignore rectangles or perform other checks


def update_presentation(df, presentation, update_count, entries_per_slide, vertical_movement_per_entry, event):
    participant_count = df.shape[0]  # Total number of new participants to ad
    entries_per_row = df.shape[1]  # Assuming this is used somewhere in populate_group
    initial_slide_index = 3 # Start from this slide index
    horizontal_movement_per_entry = -944  # Horizontal movement for each entry
    TRANSITION_IN_SECONDS = 3.5

    slide = presentation.Slides(initial_slide_index)
    group_objects = collect_group_shapes(slide)

    for row_index in range(participant_count):
        if update_count % entries_per_slide == 0 and update_count != 0 and update_count != entries_per_slide:
            logging.debug(f"Current update_count: {update_count}, triggering slide duplication.")
            duplicated_slide = slide.Duplicate().Item(1)
            slide = duplicated_slide
            group_objects = collect_group_shapes(slide)

            for group in group_objects[1:]:
                group.Top -= vertical_movement_per_entry * entries_per_slide

            logging.info('Duplicated slide after %s participants', update_count)
            
            event.wait(1)

            logging.info('Going to the next slide, total slides %s', presentation.Slides.Count)
            assert presentation.SlideShowWindow, 'no active slideshow'
            presentation.SlideShowWindow.View.Next()
            event.wait(2*entries_per_slide + 2*TRANSITION_IN_SECONDS)

        group_objects[1].Copy()
        pasted_group = slide.Shapes.Paste()
        pasted_group.ZOrder(1)
        vertical_adjustment = vertical_movement_per_entry * update_count
        horizontal_adjustment = horizontal_movement_per_entry
        row = df.iloc[row_index].tolist()
        populate_group(pasted_group, row)
        pasted_group.Top = group_objects[1].Top + vertical_adjustment
        pasted_group.Left = group_objects[1].Left + horizontal_adjustment

        event.wait(.75)

        update_count += 1   

    if update_count % entries_per_slide == 0 and update_count != 0 and update_count != entries_per_slide:
        logging.info('Adding another slide after %s participants', update_count)
        duplicated_slide = slide.Duplicate().Item(1)
        slide = duplicated_slide
        group_objects = collect_group_shapes(slide)
        
        for group in group_objects[1:]:
            group.Top -= vertical_movement_per_entry * entries_per_slide
        
        event.wait(1)
        logging.info('Going to the next slide, total slides %s', presentation.Slides.Count)
        assert presentation.SlideShowWindow, 'no active slideshow'
        presentation.SlideShowWindow.View.Next()
        event.wait(2*entries_per_slide + 2*TRANSITION_IN_SECONDS)       

    return update_count