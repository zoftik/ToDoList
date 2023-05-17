import React from "react"; /// can import directly as component instead of react
import img from "./assets/avengers.jpeg";

class MovieCard extends React.Component {
  render() {
    return (
      <div className="main">
        <div className="movie-card">
          <div className="left">
            <img alt="poster" src={img} />
          </div>
          <div className="right">
            <div className="title">The Avengers</div>
            <div className="plot">It is supernatural power shown in movie</div>
            <div className="price">Rs.199</div>

            <div className="footer">
              <div className="rating">8.9</div>
              <div className="star-dis">
                <img className="str-btn" alt="decrease" src="https://cdn-icons-png.flaticon.com/128/56/56889.png" />
                <img
                  alt="star"
                  src="https://cdn-icons-png.flaticon.com/128/2107/2107957.png"
                  className="stars"
                />
                <img className="str-btn" alt="increase" src="https://cdn-icons-png.flaticon.com/128/3524/3524388.png" />
                <span>0</span>
              </div>
              <button className="favourite-btn">Favorite</button>
              <button className="cart-btn">Add To Cart</button>
            </div>
          </div>
        </div>
      </div>
    );
  }
}

export default MovieCard;
