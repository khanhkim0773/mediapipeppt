import cv2
import mediapipe as mp
import win32com.client
import time
from PIL import Image, ImageDraw, ImageFont
import numpy as np

# Khởi tạo các thành phần của MediaPipe
mp_hands = mp.solutions.hands
mp_drawing = mp.solutions.drawing_utils
hands = mp_hands.Hands(min_detection_confidence=0.8, min_tracking_confidence=0.8)

# Cài đặt camera và kích thước hiển thị
cap = cv2.VideoCapture(0)
desired_width, desired_height = 800, 600

# Trạng thái trình chiếu
is_presentation_active = False
last_gesture_time = last_slide_change_time = time.time()
previous_x = None

# Cài đặt thời gian hiển thị cử chỉ
gesture_display_duration = 3
last_displayed_gesture = ""
last_display_update_time = time.time()

# Mở PowerPoint và trình chiếu
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = True  
presentation = powerpoint.Presentations.Open(r"D:\\QNU\\NAM_4\\HK1\\THTTNT\\code\\project\\slide.ppt")

font_path = "arial.ttf"
font = ImageFont.truetype(font_path, 32)

def start_presentation():
    global is_presentation_active
    if not is_presentation_active:
        try:
            presentation.SlideShowSettings.Run()
            is_presentation_active = True
            print("Mở chế độ trình chiếu")
        except Exception as e:
            print(f"Error starting presentation: {e}")

def stop_presentation():
    global is_presentation_active
    if is_presentation_active:
        try:
            if presentation.SlideShowWindow:
                presentation.SlideShowWindow.View.Exit()
                is_presentation_active = False
                print("Đóng chế độ trình chiếu")
        except Exception as e:
            print(f"Error stopping presentation: {e}")

def next_slide():
    try:
        if is_presentation_active and presentation.SlideShowWindow:
            presentation.SlideShowWindow.View.Next()
            print("Slide sau")
    except Exception as e:
        print(f"Error moving to next slide: {e}")

def previous_slide():
    try:
        if is_presentation_active and presentation.SlideShowWindow:
            presentation.SlideShowWindow.View.Previous()
            print("Slide trước")
    except Exception as e:
        print(f"Error moving to previous slide: {e}")

def close_powerpoint():
    try:
        presentation.Close()
        powerpoint.Quit()
        print("Đã tắt PowerPoint")
    except Exception as e:
        print(f"Error closing PowerPoint: {e}")

# Nhận diện cử chỉ nắm đấm
def is_fist_closed(hand_landmarks):
    wrist = hand_landmarks.landmark[mp_hands.HandLandmark.WRIST]
    is_close_to_wrist = all(
        ((hand_landmarks.landmark[tip].x - wrist.x) ** 2 + (hand_landmarks.landmark[tip].y - wrist.y) ** 2) ** 0.5 < 0.25
        for tip in [
            mp_hands.HandLandmark.THUMB_TIP,
            mp_hands.HandLandmark.INDEX_FINGER_TIP,
            mp_hands.HandLandmark.MIDDLE_FINGER_TIP,
            mp_hands.HandLandmark.RING_FINGER_TIP,
            mp_hands.HandLandmark.PINKY_TIP
        ]
    )
    return is_close_to_wrist

# Nhận diện bàn tay mở
def is_hand_open(hand_landmarks):
    return all(
        hand_landmarks.landmark[tip].y < hand_landmarks.landmark[dip].y
        for tip, dip in zip(
            [mp_hands.HandLandmark.THUMB_TIP, mp_hands.HandLandmark.INDEX_FINGER_TIP,
             mp_hands.HandLandmark.MIDDLE_FINGER_TIP, mp_hands.HandLandmark.RING_FINGER_TIP,
             mp_hands.HandLandmark.PINKY_TIP],
            [mp_hands.HandLandmark.THUMB_IP, mp_hands.HandLandmark.INDEX_FINGER_DIP,
             mp_hands.HandLandmark.MIDDLE_FINGER_DIP, mp_hands.HandLandmark.RING_FINGER_DIP,
             mp_hands.HandLandmark.PINKY_DIP]
        )
    )

# Nhận diện cử chỉ hai ngón (Victory - chữ V)
def is_victory_gesture(hand_landmarks):
    index_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
    middle_tip = hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_TIP]
    ring_tip = hand_landmarks.landmark[mp_hands.HandLandmark.RING_FINGER_TIP]
    pinky_tip = hand_landmarks.landmark[mp_hands.HandLandmark.PINKY_TIP]
    
    return (
        index_tip.y < hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_DIP].y and
        middle_tip.y < hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_DIP].y and
        ring_tip.y > hand_landmarks.landmark[mp_hands.HandLandmark.RING_FINGER_DIP].y and
        pinky_tip.y > hand_landmarks.landmark[mp_hands.HandLandmark.PINKY_DIP].y and
        abs(index_tip.x - middle_tip.x) > 0.1  # Khoảng cách giữa hai ngón tay đủ lớn để phân biệt
    )

# Nhận diện cử chỉ ba ngón tay giơ lên
def is_three_fingers(hand_landmarks):
    index_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
    middle_tip = hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_TIP]
    ring_tip = hand_landmarks.landmark[mp_hands.HandLandmark.RING_FINGER_TIP]
    pinky_tip = hand_landmarks.landmark[mp_hands.HandLandmark.PINKY_TIP]
    
    return (
        index_tip.y < hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_DIP].y and
        middle_tip.y < hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_DIP].y and
        ring_tip.y < hand_landmarks.landmark[mp_hands.HandLandmark.RING_FINGER_DIP].y and
        pinky_tip.y > hand_landmarks.landmark[mp_hands.HandLandmark.PINKY_DIP].y  # Chỉ ngón út hướng xuống
    )


# Nhận diện chỉ ngón tay (chỉ trỏ)
def is_pointing(hand_landmarks):
    return (
        hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP].y < 
        hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_PIP].y and
        hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_TIP].y > 
        hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_PIP].y
    )

# Nhận diện cử chỉ ngón cái chỉ lên
def is_thumb_up(hand_landmarks):
    thumb_tip = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP]
    wrist = hand_landmarks.landmark[mp_hands.HandLandmark.WRIST]
    
    is_thumb_up_gesture = (
        thumb_tip.y < wrist.y and
        all(
            ((hand_landmarks.landmark[tip].x - wrist.x) ** 2 + (hand_landmarks.landmark[tip].y - wrist.y) ** 2) ** 0.5 < 0.25
            for tip in [
                mp_hands.HandLandmark.INDEX_FINGER_TIP,
                mp_hands.HandLandmark.MIDDLE_FINGER_TIP,
                mp_hands.HandLandmark.RING_FINGER_TIP,
                mp_hands.HandLandmark.PINKY_TIP
            ]
        )
    )
    return is_thumb_up_gesture

def go_to_first_slide():
    try:
        if is_presentation_active and presentation.SlideShowWindow:
            presentation.SlideShowWindow.View.GotoSlide(1)
            print("Chuyển đến slide đầu tiên")
    except Exception as e:
        print(f"Error going to first slide: {e}")

def go_to_last_slide():
    try:
        if is_presentation_active and presentation.SlideShowWindow:
            last_slide_index = presentation.Slides.Count
            presentation.SlideShowWindow.View.GotoSlide(last_slide_index)
            print("Chuyển đến slide cuối cùng")
    except Exception as e:
        print(f"Error going to last slide: {e}")

# # Xử lý cử chỉ
# def process_hand_gesture(hand_landmarks, current_time):
#     global is_presentation_active, last_gesture_time, previous_x, last_slide_change_time
#     gesture_text = ""

#     # Điều khiển chế độ trình chiếu và tắt PowerPoint
#     if current_time - last_gesture_time > 2:
#         if is_fist_closed(hand_landmarks):  # Cử chỉ "nắm đấm" để thoát trình chiếu
#             if is_presentation_active:
#                 stop_presentation()  # Chỉ dừng trình chiếu, không tắt PowerPoint
#                 gesture_text = "Dừng trình chiếu"
#             last_gesture_time = current_time
#         elif is_hand_open(hand_landmarks):
#             start_presentation()
#             last_gesture_time = current_time
#             gesture_text = "Mở chế độ trình chiếu"
#         elif is_victory_gesture(hand_landmarks):
#             go_to_first_slide()
#             last_gesture_time = current_time
#             gesture_text = "Chuyển đến slide đầu tiên"
#         elif is_three_fingers(hand_landmarks):
#             go_to_last_slide()
#             last_gesture_time = current_time
#             gesture_text = "Chuyển đến slide cuối cùng"
#         elif is_thumb_up(hand_landmarks):
#             close_powerpoint()  # Chỉ gọi hàm này khi thực sự muốn tắt PowerPoint
#             last_gesture_time = current_time
#             gesture_text = "Tắt PowerPoint"

#     # Điều khiển slide
#     if is_presentation_active:
#         index_finger_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
#         index_finger_x = int(index_finger_tip.x * desired_width)
        
#         if is_pointing(hand_landmarks):
#             if previous_x is not None and current_time - last_slide_change_time > 1:
#                 if index_finger_x < previous_x - 50:
#                     previous_slide()
#                     last_slide_change_time = current_time
#                     gesture_text = "Slide trước"
#                 elif index_finger_x > previous_x + 50:
#                     next_slide()
#                     last_slide_change_time = current_time
#                     gesture_text = "Slide sau"
#             previous_x = index_finger_x
    
#     return gesture_text


def process_hand_gesture(hand_landmarks, current_time):
    global is_presentation_active, last_gesture_time, previous_x, last_slide_change_time
    gesture_text = ""

    # Điều khiển chế độ trình chiếu và tắt PowerPoint
    if current_time - last_gesture_time > 2:
        if is_fist_closed(hand_landmarks):  # Cử chỉ "nắm đấm" để thoát trình chiếu
            stop_presentation()  # Chỉ dừng trình chiếu, không tắt PowerPoint
            gesture_text = "Dừng trình chiếu"
            last_gesture_time = current_time
        elif is_hand_open(hand_landmarks):
            start_presentation()
            last_gesture_time = current_time
            gesture_text = "Mở chế độ trình chiếu"
        elif is_victory_gesture(hand_landmarks):
            go_to_first_slide()
            last_gesture_time = current_time
            gesture_text = "Chuyển đến slide đầu tiên"
        elif is_three_fingers(hand_landmarks):
            go_to_last_slide()
            last_gesture_time = current_time
            gesture_text = "Chuyển đến slide cuối cùng"
        elif is_thumb_up(hand_landmarks):
            close_powerpoint()  # Chỉ gọi hàm này khi thực sự muốn tắt PowerPoint
            last_gesture_time = current_time
            gesture_text = "Tắt PowerPoint"

    # Điều khiển slide
    if is_presentation_active:
        index_finger_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
        index_finger_x = int(index_finger_tip.x * desired_width)
        
        if is_pointing(hand_landmarks):
            if previous_x is not None and current_time - last_slide_change_time > 1:
                if index_finger_x < previous_x - 50:
                    previous_slide()
                    last_slide_change_time = current_time
                    gesture_text = "Slide trước"
                elif index_finger_x > previous_x + 50:
                    next_slide()
                    last_slide_change_time = current_time
                    gesture_text = "Slide sau"
            previous_x = index_finger_x
    
    return gesture_text

# Hiển thị văn bản tiếng Việt trên camera
def display_text(image, text, position):
    pil_image = Image.fromarray(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(pil_image)
    draw.text(position, text, font=font, fill=(255, 0, 0))
    return cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)

# Main loop
while cap.isOpened():
    success, image = cap.read()
    if not success:
        break
    
    image = cv2.cvtColor(cv2.resize(image, (desired_width, desired_height)), cv2.COLOR_BGR2RGB)
    results = hands.process(image)
    image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
    
    current_time = time.time()
    gesture_text = ""
    
    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            mp_drawing.draw_landmarks(image, hand_landmarks, mp_hands.HAND_CONNECTIONS)
            detected_gesture = process_hand_gesture(hand_landmarks, current_time)
            
            if detected_gesture and (detected_gesture != last_displayed_gesture or (current_time - last_display_update_time) > gesture_display_duration):
                last_displayed_gesture = detected_gesture
                last_display_update_time = current_time
            gesture_text = last_displayed_gesture if (current_time - last_display_update_time) <= gesture_display_duration else ""
    
    image = display_text(image, f'{gesture_text}', (10, 40))
    
    cv2.imshow('MediaPipe Hands', image)
    
    if cv2.waitKey(5) & 0xFF == 27:
        break

cap.release()
cv2.destroyAllWindows()
presentation.Close()
