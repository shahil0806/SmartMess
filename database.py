from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()


class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    student_uid = db.Column(
        db.String(50),
        unique=True,
        nullable=False
    )

    name = db.Column(
        db.String(100),
        nullable=False
    )

    roll_number = db.Column(
        db.String(50),
        unique=True,
        nullable=False
    )

    branch = db.Column(
        db.String(100),
        nullable=False
    )

    hostel_room = db.Column(
        db.String(50),
        nullable=False
    )

    photo = db.Column(
        db.String(255),
        nullable=True
    )

    is_active = db.Column(
        db.Boolean,
        default=True
    )

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )


class MealRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    student_uid = db.Column(
        db.String(50),
        nullable=False
    )

    meal = db.Column(
        db.String(20),
        nullable=False
    )

    coupon_token = db.Column(
        db.String(100),
        unique=True,
        nullable=False
    )

    generated_at = db.Column(
        db.DateTime,
        nullable=False
    )

    expires_at = db.Column(
        db.DateTime,
        nullable=False
    )

    used_at = db.Column(
        db.DateTime,
        nullable=True
    )

    status = db.Column(
        db.String(20),
        default="GENERATED"
    )