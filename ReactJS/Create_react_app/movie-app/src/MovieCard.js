import React from "react"; /// can import directly as component instead of react

class MovieCard extends React.Component {
  render() {
    return (
      <div className="main">
        <div className="movie-card">
          <div className="left">
            <img alt="poster" />
          </div>
          <div className="right">
            <div className="title">The Avengers</div>
            <div className="plot">It is supernatural power shown in movie</div>
            <div className="price">Rs.199</div>

            <div className="footer">
              <div className="rating">8.9</div>
              <div className="stars">star</div>
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
