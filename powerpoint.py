import slide_making as sm
from pptx import Presentation


def main():
    template_file = 'template.pptx'
    prs = Presentation(template_file)
    for i in range(1, 23):
        try:
            print(i)
            sm.slide_writer(i, prs)
        except Exception:
            print(f"Error on slide {i}")

    prs.save('example_from_template.pptx')


if __name__ == "__main__":
    main()
