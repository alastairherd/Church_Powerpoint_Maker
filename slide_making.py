from functions import *

class SlideFill:
    def __init__(self, slide):
        self.slide = slide
    
    def fill_main(self, title, body, copy):
        for idx, placeholder in enumerate(self.slide.placeholders):
            try:
                placeholder.text_frame.clear()
                if idx == 0:
                    new_body = title
                elif idx == 1:
                    new_body = body
                elif idx == 2:
                    new_body = copy
                placeholder.text_frame.text = new_body
            except:
                pass
    
    def fill_component(self, title, body, address):
        for idx, placeholder in enumerate(self.slide.placeholders):
            try:
                placeholder.text_frame.clear()
                if idx == 0:
                    placeholder.text_frame.text = title
                elif idx == 1:
                    p = placeholder.text_frame.add_paragraph()
                    run = p.add_run()
                    run.font.size = Pt(28)
                    run.font.color.rgb = RGBColor(161, 38, 38)
                    run.text = f"{address} "
                    run = p.add_run()
                    run.font.size = Pt(28)
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    run.text = body    
            except:
                pass
    
    def fill_psalm(self, title, body, copy, meter):
        for idx, placeholder in enumerate(self.slide.placeholders):
            try:
                placeholder.text_frame.clear()
                if idx == 0:
                    new_body = title
                elif idx == 1:
                    new_body = body
                elif idx == 2:
                    new_body = copy
                elif idx == 3:
                    new_body = meter
                placeholder.text_frame.text = new_body
            except:
                pass

    def fill_reading(self,title,body):
        for idx, placeholder in enumerate(self.slide.placeholders):
            try:
                placeholder.text_frame.clear()
                if idx == 0:
                    new_body = title
                elif idx == 1:
                    new_body = body
                placeholder.text_frame.text = new_body
            except:
                pass

############################################################################################################

def flag_format(flag):
    try:
        ## Such as 'notices' and 'song1'
        if type(flag) == str:
            return globals()[flag]
        elif type(flag) == list:
            try:
                ## For 'component versions
                if type(flag) == list:
                    return component_assigner(flag[0])
            except:
                ## For 'call to worship'
                return (globals()[flag[0]],flag[1])
    except:
        return "Not found"

def slide_maker(layout_type, prs):
    slide_layout = prs.slide_layouts[layout_type]
    return prs.slides.add_slide(slide_layout)


def slide_writer(flag, prs):
    # Populate the placeholders on the slide with data from variables
    flag_val = flag_format(slide_dict[flag])

    
    if flag == 3:
        slide = slide_maker(1, prs)
        # This works for call to worship
        title = flag_val[1]
        body = flag_val[0][0]
        copy = flag_val[0][1]
        SlideFill(slide).fill_main(title, body, copy)
    
    elif slide_dict[flag] == "goodbye":
        slide = slide_maker(7, prs)

    elif flag in song_list and flag_val != "Not found":
        # This works for psalms
        try:
            flag_val[0].find("Psalm") != -1
            list_len = len(flag_val[2])
            for i in range(0,list_len):
                slide = slide_maker(3, prs)
                title = flag_val[0]
                body = flag_val[2][i]
                copy = f"Words: Sing Psalms! © 2003 Free Church of Scotland\nComposer: {psalm_tune[0]}\nTune: {psalm_tune[1]}\n©: Public Domain\nCCLI: 522221"
                meter = f"Meter: {flag_val[1]}"
                SlideFill(slide).fill_psalm(title, body, copy, meter)
        except:
            # This works for songs
            try:
                list_len = len(flag_val[0])
                for i in range(0,list_len):
                    slide = slide_maker(0, prs)
                    title = flag_val[1]
                    body = flag_val[0][i]
                    # Could put this outside the loop, in the future to only fill on the final slide
                    copy = f"Words: {flag_val[2]}\nComposer: {flag_val[3]}\nTune: {flag_val[4]}\n©: {flag_val[5]}\nCCLI: 522221"
                    SlideFill(slide).fill_main(title, body, copy)
            except:
                slide = slide_maker(0, prs)
                try:
                    title = flag_val[1]
                except:
                    title = slide_dict[flag]
                body = "Not in Public Domain"
                try:
                    copy = f"Words: {flag_val[2]}\nComposer: {flag_val[3]}\nTune: {flag_val[4]}\n©: {flag_val[5]}\nCCLI: 522221"
                except:
                    copy = "Error"
                SlideFill(slide).fill_main(title, body, copy)
                
    elif flag in component_list:
        # This works for components
        list_len = len(flag_val[1])
        for i in range(0,list_len):
            slide = slide_maker(2, prs)
            title = slide_dict[flag][1]
            address = flag_val[0]
            body = flag_val[1][i]
            SlideFill(slide).fill_component(title, body, address)
    
    elif flag in reading_list:
        title = flag_val[1]
        body = f"{flag_val[0]}\n\npg. X"
        slide = slide_maker(4, prs)
        SlideFill(slide).fill_reading(title, body)

    elif flag in catechism_list:
        title = f"Westminster Shorter Catechism {flag_val[2]}"
        question = f"\n\n{flag_val[1]}"
        answer = f"{flag_val[0]}"
        slide = slide_maker(5, prs)
        SlideFill(slide).fill_component(title, question, answer)

    elif "prayer" in slide_dict[flag][0]:
        title = slide_dict[flag][1]
        body = ""
        slide = slide_maker(4, prs)
        SlideFill(slide).fill_reading(title, body)